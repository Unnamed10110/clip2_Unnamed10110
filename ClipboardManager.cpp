#include "ClipboardManager.h"
#include <sstream>
#include <iomanip>
#include <ctime>
#include <fstream>
#include <algorithm>
#include <cmath>
#include <cwctype>
#include <cstring>
#include <map>
#include <shlobj.h>
#include <mmsystem.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <objbase.h>
#include <oleauto.h>
#include <richedit.h>
#include <wincrypt.h>
#include <dpapi.h>
#pragma comment(lib, "crypt32.lib")
#ifdef HAVE_UIAUTOMATION
#include <UIAutomation.h>
#pragma comment(lib, "UIAutomationCore.lib")
#endif
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mf.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")
#ifndef MSFTEDIT_CLASS
#define MSFTEDIT_CLASS L"RichEdit50W"
#endif
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")

// AMOLED neon palette. Background stays pure black (true OLED), the rest of the
// colors swap depending on the active theme preset selected at runtime.
namespace Theme5250 {
    COLORREF BG        = RGB(0, 0, 0);
    COLORREF TXT       = RGB(0, 255, 102);
    COLORREF SEL_BG    = RGB(0, 255, 102);
    COLORREF SEL_FG    = RGB(0, 0, 0);
    COLORREF BORDER    = RGB(0, 200, 80);
    COLORREF DIM       = RGB(0, 128, 40);
}

enum ThemeId {
    THEME_NEON_GREEN  = 0,
    THEME_NEON_RED    = 1,
    THEME_NEON_BLUE   = 2,
    THEME_NEON_CYAN   = 3,
    THEME_NEON_PURPLE = 4,
    THEME_NEON_YELLOW = 5,
    THEME_NEON_ORANGE = 6,
    THEME_NEON_WHITE  = 7,
    THEME_COUNT
};

struct ThemePreset {
    const wchar_t* name;
    COLORREF txt;
    COLORREF selBg;
    COLORREF border;
    COLORREF dim;
};

static const ThemePreset kThemePresets[THEME_COUNT] = {
    { L"Neon Green (AS/400)", RGB(0, 255, 102),   RGB(0, 255, 102),   RGB(0, 200, 80),    RGB(0, 128, 40) },
    { L"Neon Red",            RGB(255, 32, 64),   RGB(255, 32, 64),   RGB(220, 20, 50),   RGB(120, 8, 24) },
    { L"Neon Blue",           RGB(40, 120, 255),  RGB(40, 120, 255),  RGB(30, 100, 220),  RGB(16, 56, 140) },
    { L"Neon Cyan",           RGB(0, 240, 255),   RGB(0, 240, 255),   RGB(0, 200, 230),   RGB(0, 100, 120) },
    { L"Neon Purple",         RGB(200, 60, 255),  RGB(200, 60, 255),  RGB(170, 40, 220),  RGB(90, 20, 130) },
    { L"Neon Yellow",         RGB(255, 230, 0),   RGB(255, 230, 0),   RGB(220, 200, 0),   RGB(120, 110, 0) },
    { L"Neon Orange",         RGB(255, 128, 0),   RGB(255, 128, 0),   RGB(220, 110, 0),   RGB(130, 60, 0) },
    { L"Neon White",          RGB(240, 240, 240), RGB(240, 240, 240), RGB(200, 200, 200), RGB(110, 110, 110) },
};

static int g_currentThemeId = THEME_NEON_GREEN;

// Sentinel meaning "no user override — follow the active preset's color".
// COLORREF is 0x00BBGGRR; a high byte of 0xFF is invalid and safe as a marker.
static const COLORREF kThemeFontColorPreset = (COLORREF)0xFFFFFFFF;
static COLORREF g_themeFontColor = kThemeFontColorPreset;

// Overlay/search font family. Empty means use the default. Default font keeps the
// terminal aesthetic of the original 5250 theme.
static const wchar_t* kDefaultThemeFontFace = L"Consolas";
static std::wstring g_themeFontFace = kDefaultThemeFontFace;

// ---- Per-element color overrides -------------------------------------------------
// Each overlay color can be individually overridden by the user (Settings -> Colors).
// A slot left at kColorUsePreset follows the active theme preset (or the legacy accent
// override) so existing themes keep working untouched.
enum ThemeColorSlot {
    TC_BG = 0,     // window background
    TC_TXT,        // primary text
    TC_SELBG,      // selection / accent background
    TC_SELFG,      // text drawn on the selection
    TC_BORDER,     // borders, badges, pills, scrollbar thumb
    TC_DIM,        // separators / secondary chrome
    TC_COUNT
};
static const COLORREF kColorUsePreset = (COLORREF)0xFF000000;  // high byte invalid => sentinel
static COLORREF g_colorOverride[TC_COUNT] = {
    kColorUsePreset, kColorUsePreset, kColorUsePreset,
    kColorUsePreset, kColorUsePreset, kColorUsePreset
};
static const wchar_t* const kColorRegNames[TC_COUNT] = {
    L"ColorBG", L"ColorTXT", L"ColorSELBG", L"ColorSELFG", L"ColorBORDER", L"ColorDIM"
};
static const wchar_t* const kColorLabels[TC_COUNT] = {
    L"BG", L"Text", L"Accent", L"Sel", L"Border", L"Dim"
};

static void ApplyThemeId(int id) {
    if (id < 0 || id >= THEME_COUNT) id = THEME_NEON_GREEN;
    const ThemePreset& p = kThemePresets[id];
    g_currentThemeId = id;

    // 1. Start from the preset.
    Theme5250::BG     = RGB(0, 0, 0);
    Theme5250::SEL_FG = RGB(0, 0, 0);
    Theme5250::BORDER = p.border;
    Theme5250::DIM    = p.dim;
    Theme5250::TXT    = p.txt;
    Theme5250::SEL_BG = p.selBg;

    // 2. Legacy single "Font color" accent override (kept for backward compatibility):
    //    when set, it recolors both text and selection background.
    if (g_themeFontColor != kThemeFontColorPreset) {
        Theme5250::TXT    = g_themeFontColor;
        Theme5250::SEL_BG = g_themeFontColor;
    }

    // 3. Granular per-element overrides take final precedence.
    if (g_colorOverride[TC_BG]     != kColorUsePreset) Theme5250::BG     = g_colorOverride[TC_BG];
    if (g_colorOverride[TC_TXT]    != kColorUsePreset) Theme5250::TXT    = g_colorOverride[TC_TXT];
    if (g_colorOverride[TC_SELBG]  != kColorUsePreset) Theme5250::SEL_BG = g_colorOverride[TC_SELBG];
    if (g_colorOverride[TC_SELFG]  != kColorUsePreset) Theme5250::SEL_FG = g_colorOverride[TC_SELFG];
    if (g_colorOverride[TC_BORDER] != kColorUsePreset) Theme5250::BORDER = g_colorOverride[TC_BORDER];
    if (g_colorOverride[TC_DIM]    != kColorUsePreset) Theme5250::DIM    = g_colorOverride[TC_DIM];
}

// Current effective color for a slot (after preset + overrides have been applied).
static COLORREF EffectiveThemeColor(int slot) {
    switch (slot) {
        case TC_BG:     return Theme5250::BG;
        case TC_TXT:    return Theme5250::TXT;
        case TC_SELBG:  return Theme5250::SEL_BG;
        case TC_SELFG:  return Theme5250::SEL_FG;
        case TC_BORDER: return Theme5250::BORDER;
        case TC_DIM:    return Theme5250::DIM;
        default:        return Theme5250::TXT;
    }
}

static int LoadThemeIdFromRegistry() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return THEME_NEON_GREEN;
    DWORD val = THEME_NEON_GREEN, sz = sizeof(DWORD), type = REG_DWORD;
    if (RegQueryValueExW(hKey, L"ThemeId", nullptr, &type, (LPBYTE)&val, &sz) != ERROR_SUCCESS || type != REG_DWORD)
        val = THEME_NEON_GREEN;
    RegCloseKey(hKey);
    if (val >= THEME_COUNT) val = THEME_NEON_GREEN;
    return (int)val;
}

static void SaveThemeIdToRegistry(int id) {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    DWORD v = (DWORD)id;
    RegSetValueExW(hKey, L"ThemeId", 0, REG_DWORD, (BYTE*)&v, sizeof(DWORD));
    RegCloseKey(hKey);
}

static COLORREF LoadThemeFontColorFromRegistry() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return kThemeFontColorPreset;
    DWORD val = kThemeFontColorPreset, sz = sizeof(DWORD), type = REG_DWORD;
    LONG rc = RegQueryValueExW(hKey, L"ThemeFontColor", nullptr, &type, (LPBYTE)&val, &sz);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_DWORD) return kThemeFontColorPreset;
    return (COLORREF)val;
}

static void SaveThemeFontColorToRegistry(COLORREF c) {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    if (c == kThemeFontColorPreset) {
        // Clear the override entirely so the preset color takes back over on next launch.
        RegDeleteValueW(hKey, L"ThemeFontColor");
    } else {
        DWORD v = (DWORD)c;
        RegSetValueExW(hKey, L"ThemeFontColor", 0, REG_DWORD, (BYTE*)&v, sizeof(DWORD));
    }
    RegCloseKey(hKey);
}

static void LoadThemeColorsFromRegistry() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
        for (int i = 0; i < TC_COUNT; i++) g_colorOverride[i] = kColorUsePreset;
        return;
    }
    for (int i = 0; i < TC_COUNT; i++) {
        DWORD val = 0, sz = sizeof(DWORD), type = REG_DWORD;
        if (RegQueryValueExW(hKey, kColorRegNames[i], nullptr, &type, (LPBYTE)&val, &sz) == ERROR_SUCCESS && type == REG_DWORD)
            g_colorOverride[i] = (COLORREF)val;
        else
            g_colorOverride[i] = kColorUsePreset;
    }
    RegCloseKey(hKey);
}

static void SaveThemeColorToRegistry(int slot, COLORREF c) {
    if (slot < 0 || slot >= TC_COUNT) return;
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    if (c == kColorUsePreset) {
        RegDeleteValueW(hKey, kColorRegNames[slot]);  // clear override -> follow preset
    } else {
        DWORD v = (DWORD)c;
        RegSetValueExW(hKey, kColorRegNames[slot], 0, REG_DWORD, (BYTE*)&v, sizeof(DWORD));
    }
    RegCloseKey(hKey);
}

static std::wstring LoadThemeFontFaceFromRegistry() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return kDefaultThemeFontFace;
    wchar_t buf[LF_FACESIZE + 1] = {};
    DWORD bytes = sizeof(buf) - sizeof(wchar_t);
    DWORD type = REG_SZ;
    LONG rc = RegQueryValueExW(hKey, L"ThemeFontFace", nullptr, &type, (LPBYTE)buf, &bytes);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_SZ || buf[0] == 0) return kDefaultThemeFontFace;
    return std::wstring(buf);
}

static void SaveThemeFontFaceToRegistry(const std::wstring& face) {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    if (face.empty() || face == kDefaultThemeFontFace) {
        RegDeleteValueW(hKey, L"ThemeFontFace");
    } else {
        std::wstring clipped = face.size() <= LF_FACESIZE - 1 ? face : face.substr(0, LF_FACESIZE - 1);
        RegSetValueExW(hKey, L"ThemeFontFace", 0, REG_SZ,
                       (const BYTE*)clipped.c_str(), (DWORD)((clipped.size() + 1) * sizeof(wchar_t)));
    }
    RegCloseKey(hKey);
}

static int LoadMaxItemsFromRegistry(int defaultValue, int minValue, int maxValue) {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return defaultValue;
    DWORD val = (DWORD)defaultValue, sz = sizeof(DWORD), type = REG_DWORD;
    LONG rc = RegQueryValueExW(hKey, L"MaxItems", nullptr, &type, (LPBYTE)&val, &sz);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_DWORD) return defaultValue;
    int v = (int)val;
    if (v < minValue) v = minValue;
    if (v > maxValue) v = maxValue;
    return v;
}

static void SaveMaxItemsToRegistry(int value) {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    DWORD v = (DWORD)value;
    RegSetValueExW(hKey, L"MaxItems", 0, REG_DWORD, (BYTE*)&v, sizeof(DWORD));
    RegCloseKey(hKey);
}

// Case-insensitive fuzzy match with ranking. Returns a positive score on match, 0 on no match.
// Exact substring matches score highest; subsequence matches score by contiguity and word-start bonuses.
static int FuzzyScore(const std::wstring& needleLower, const std::wstring& haystack) {
    if (needleLower.empty()) return 1;
    if (haystack.empty()) return 0;
    std::wstring hay = haystack;
    std::transform(hay.begin(), hay.end(), hay.begin(), ::towlower);

    // Exact substring is the strongest signal.
    size_t pos = hay.find(needleLower);
    if (pos != std::wstring::npos) {
        int score = 10000;
        if (pos == 0) score += 500;                              // prefix match
        else if (hay[pos - 1] == L' ' || hay[pos - 1] == L'\n' ||
                 hay[pos - 1] == L'\t' || hay[pos - 1] == L'/' ||
                 hay[pos - 1] == L'.')
            score += 250;                                        // word-start match
        score -= (int)std::min(pos, (size_t)200);                // earlier is better
        return score;
    }

    // Subsequence match: every needle char must appear in order.
    int score = 0;
    size_t h = 0;
    bool prevMatched = false;
    bool atWordStart = true;
    for (size_t n = 0; n < needleLower.size(); n++) {
        wchar_t want = needleLower[n];
        bool found = false;
        for (; h < hay.size(); h++) {
            wchar_t c = hay[h];
            bool wordStart = (h == 0) || hay[h - 1] == L' ' || hay[h - 1] == L'\n' ||
                             hay[h - 1] == L'\t' || hay[h - 1] == L'/' || hay[h - 1] == L'.';
            if (c == want) {
                score += 10;
                if (prevMatched) score += 15;   // contiguous run bonus
                if (wordStart) score += 20;      // word-start bonus
                prevMatched = true;
                h++;
                found = true;
                break;
            } else {
                prevMatched = false;
            }
            (void)atWordStart;
        }
        if (!found) return 0;
    }
    return score;
}

// Shared font handle used by the overlay paint loop and the search box. Re-created on
// demand whenever the user picks a different face from Settings. The previous code
// created two separate hard-coded "Consolas" fonts (one in CreateWindowEx setup for
// the search edit, one inside WM_PAINT); they are unified here.
static HFONT g_overlayFont = nullptr;
static std::wstring g_overlayFontCachedFace;

static HFONT GetOverlayFont() {
    const std::wstring& want = g_themeFontFace.empty() ? std::wstring(kDefaultThemeFontFace) : g_themeFontFace;
    if (g_overlayFont != nullptr && g_overlayFontCachedFace == want)
        return g_overlayFont;
    if (g_overlayFont) {
        DeleteObject(g_overlayFont);
        g_overlayFont = nullptr;
    }
    g_overlayFont = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, want.c_str());
    g_overlayFontCachedFace = want;
    return g_overlayFont;
}

// Bold + small variants of the overlay font, used for the header chrome and metadata.
static HFONT g_overlayFontBold = nullptr;
static HFONT g_overlayFontSmall = nullptr;
static std::wstring g_overlayFontVariantFace;

static void EnsureOverlayFontVariants() {
    const std::wstring want = g_themeFontFace.empty() ? std::wstring(kDefaultThemeFontFace) : g_themeFontFace;
    if (g_overlayFontBold && g_overlayFontSmall && g_overlayFontVariantFace == want) return;
    if (g_overlayFontBold) { DeleteObject(g_overlayFontBold); g_overlayFontBold = nullptr; }
    if (g_overlayFontSmall) { DeleteObject(g_overlayFontSmall); g_overlayFontSmall = nullptr; }
    g_overlayFontBold = CreateFontW(15, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, want.c_str());
    g_overlayFontSmall = CreateFontW(12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, want.c_str());
    g_overlayFontVariantFace = want;
}
static HFONT GetOverlayFontBold()  { EnsureOverlayFontVariants(); return g_overlayFontBold; }
static HFONT GetOverlayFontSmall() { EnsureOverlayFontVariants(); return g_overlayFontSmall; }

// Linear blend between two COLORREFs (t in 0..255). Used for subtle hover/scrollbar tints.
static COLORREF BlendColor(COLORREF a, COLORREF b, int t) {
    if (t < 0) t = 0; if (t > 255) t = 255;
    int ar = GetRValue(a), ag = GetGValue(a), ab = GetBValue(a);
    int br = GetRValue(b), bg = GetGValue(b), bb = GetBValue(b);
    return RGB(ar + (br - ar) * t / 255, ag + (bg - ag) * t / 255, ab + (bb - ab) * t / 255);
}

// Compact relative timestamp ("now", "5m", "3h", "2d") for the row metadata.
static std::wstring RelativeTimeString(const std::chrono::system_clock::time_point& tp) {
    auto now = std::chrono::system_clock::now();
    auto secs = std::chrono::duration_cast<std::chrono::seconds>(now - tp).count();
    if (secs < 0) secs = 0;
    if (secs < 10) return L"now";
    if (secs < 60) return std::to_wstring(secs) + L"s";
    long long mins = secs / 60;
    if (mins < 60) return std::to_wstring(mins) + L"m";
    long long hours = mins / 60;
    if (hours < 24) return std::to_wstring(hours) + L"h";
    long long days = hours / 24;
    if (days < 7) return std::to_wstring(days) + L"d";
    long long weeks = days / 7;
    if (weeks < 5) return std::to_wstring(weeks) + L"w";
    return std::to_wstring(days / 30) + L"mo";
}

static void ResetOverlayFontCache() {
    if (g_overlayFont) {
        DeleteObject(g_overlayFont);
        g_overlayFont = nullptr;
    }
    g_overlayFontCachedFace.clear();
    if (g_overlayFontBold) { DeleteObject(g_overlayFontBold); g_overlayFontBold = nullptr; }
    if (g_overlayFontSmall) { DeleteObject(g_overlayFontSmall); g_overlayFontSmall = nullptr; }
    g_overlayFontVariantFace.clear();
}

// Best-effort check that a font family actually exists on the system before we accept
// it from the registry. Avoids ending up with a blank UI if a user-typed face is wrong.
static bool FontFamilyExists(const std::wstring& face) {
    if (face.empty()) return false;
    HDC hdc = GetDC(nullptr);
    if (!hdc) return false;
    LOGFONTW lf = {};
    wcsncpy_s(lf.lfFaceName, face.c_str(), _TRUNCATE);
    lf.lfCharSet = DEFAULT_CHARSET;
    bool found = false;
    EnumFontFamiliesExW(hdc, &lf,
        [](const LOGFONTW*, const TEXTMETRICW*, DWORD, LPARAM lParam) -> int {
            *(bool*)lParam = true;
            return 0;
        },
        (LPARAM)&found, 0);
    ReleaseDC(nullptr, hdc);
    return found;
}

ClipboardManager* ClipboardManager::instance = nullptr;

// Rich Edit stream callbacks for RTF
struct StreamCookie {
    const char* data;
    size_t pos;
    size_t size;
    std::string* outStr;
};
static DWORD CALLBACK RichEditStreamInCallback(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
    StreamCookie* c = (StreamCookie*)cookie;
    size_t remain = c->size - c->pos;
    size_t toCopy = (size_t)cb < remain ? (size_t)cb : remain;
    memcpy(buf, c->data + c->pos, toCopy);
    c->pos += toCopy;
    *pcb = (LONG)toCopy;
    return 0;
}
static DWORD CALLBACK RichEditStreamOutCallback(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
    StreamCookie* c = (StreamCookie*)cookie;
    if (c->outStr) c->outStr->append((char*)buf, (size_t)cb);
    *pcb = cb;
    return 0;
}
static bool IsRtfContent(const std::wstring& s) {
    if (s.size() < 5) return false;
    return (s[0] == L'{' && s[1] == L'\\' && (s[2] == L'r' || s[2] == L'R') && (s[3] == L't' || s[3] == L'T') && (s[4] == L'f' || s[4] == L'F'));
}

// Convert RTF bytes to a plain wide-string by streaming through a message-only RichEdit.
// Returns empty on failure. Used so that when a snippet is stored only as RTF (no
// contentPlain), we can still publish a correct CF_UNICODETEXT view of it.
static std::wstring ExtractPlainFromRtfBytes(const std::string& rtfBytes) {
    if (rtfBytes.empty()) return L"";
    static bool s_msftedit = false;
    if (!s_msftedit) {
        LoadLibraryW(L"Msftedit.dll");
        s_msftedit = true;
    }
    HWND hRe = CreateWindowExW(0, MSFTEDIT_CLASS, L"",
        ES_MULTILINE | ES_AUTOVSCROLL, 0, 0, 10, 10,
        HWND_MESSAGE, nullptr, GetModuleHandle(nullptr), nullptr);
    if (!hRe) return L"";
    StreamCookie cookie = { rtfBytes.data(), 0, rtfBytes.size(), nullptr };
    EDITSTREAM es = { (DWORD_PTR)&cookie, 0, RichEditStreamInCallback };
    SendMessageW(hRe, EM_STREAMIN, SF_RTF, (LPARAM)&es);
    LONG len = GetWindowTextLengthW(hRe);
    std::wstring out;
    if (len > 0) {
        out.resize((size_t)len + 1);
        int got = GetWindowTextW(hRe, &out[0], (int)out.size());
        if (got > 0) out.resize((size_t)got);
        else out.clear();
    }
    DestroyWindow(hRe);
    return out;
}

// ClipboardItem methods
std::wstring ClipboardItem::GetFormatName(UINT fmt) {
    if (fmt == CF_TEXT) return L"Text";
    if (fmt == CF_UNICODETEXT) return L"Unicode Text";
    if (fmt == CF_OEMTEXT) return L"OEM Text";
    if (fmt == CF_BITMAP) return L"Bitmap";
    if (fmt == CF_DIB) return L"DIB";
    if (fmt == CF_DIBV5) return L"DIB v5";
    if (fmt == CF_ENHMETAFILE) return L"Enhanced Metafile";
    if (fmt == CF_HDROP) return L"File Drop";
    if (fmt == CF_LOCALE) return L"Locale";
    if (fmt == CF_METAFILEPICT) return L"Metafile Picture";
    if (fmt == CF_PALETTE) return L"Palette";
    if (fmt == CF_PENDATA) return L"Pen Data";
    if (fmt == CF_RIFF) return L"RIFF";
    if (fmt == CF_SYLK) return L"SYLK";
    if (fmt == CF_WAVE) return L"Wave";
    if (fmt == CF_TIFF) return L"TIFF";
    
    wchar_t formatName[256] = {0};
    try {
        if (GetClipboardFormatNameW(fmt, formatName, 256)) {
            return formatName;
        }
    } catch (...) {
        // Error getting format name, return generic name
    }
    
    return L"Unknown Format";
}

std::wstring ClipboardItem::GetFileType(UINT fmt) {
    try {
        if (fmt == CF_TEXT || fmt == CF_UNICODETEXT || fmt == CF_OEMTEXT) {
            return L"Text";
        }
        if (fmt == CF_BITMAP || fmt == CF_DIB || fmt == CF_DIBV5) {
            isImage = true;
            return L"Image";
        }
        if (fmt == CF_HDROP) {
            // Check if files are images/videos
            try {
                const std::vector<BYTE>* dropData = GetFormatData(CF_HDROP);
                if (dropData && dropData->size() >= sizeof(DROPFILES)) {
                    DROPFILES* df = (DROPFILES*)dropData->data();
                    if (df) {
                        // Validate pFiles offset
                        if (df->pFiles > 0 && df->pFiles < (DWORD)dropData->size()) {
                            wchar_t* files = (wchar_t*)((BYTE*)dropData->data() + df->pFiles);
                            if (files) {
                                // Validate that we're not reading past buffer
                                size_t maxLen = (dropData->size() - df->pFiles) / sizeof(wchar_t);
                                if (maxLen > 0 && maxLen < 32767) { // Reasonable limit
                                    try {
                                        std::wstring firstFile = files;
                                        // Limit to reasonable length to prevent buffer overruns
                                        if (firstFile.length() > maxLen) {
                                            firstFile = firstFile.substr(0, maxLen);
                                        }
                                        size_t dotPos = firstFile.find_last_of(L".");
                                        if (dotPos != std::wstring::npos && dotPos < firstFile.length() - 1) {
                                            std::wstring ext = firstFile.substr(dotPos + 1);
                                            if (!ext.empty() && ext.length() < 20) {
                                                std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
                                                if (ext == L"jpg" || ext == L"jpeg" || ext == L"png" || ext == L"gif" || ext == L"bmp" || ext == L"ico" || ext == L"webp") {
                                                    isImage = true;
                                                    return L"Image";
                                                }
                                                if (ext == L"mp4" || ext == L"avi" || ext == L"mkv" || ext == L"mov" || ext == L"wmv" || ext == L"flv" || ext == L"webm") {
                                                    isVideo = true;
                                                    return L"Video";
                                                }
                                            }
                                        }
                                    } catch (...) {
                                        // Error processing filename
                                    }
                                }
                            }
                        }
                    }
                }
            } catch (...) {
                // Error processing file drop data
            }
            return L"Files";
        }
        if (fmt == CF_RIFF || fmt == CF_WAVE) {
            return L"Audio";
        }
        if (fmt == CF_ENHMETAFILE || fmt == CF_METAFILEPICT) {
            isImage = true;
            return L"Metafile";
        }
    } catch (...) {
        // Error processing file type, return generic type
    }
    return L"Other";
}

std::wstring ClipboardItem::GetPreview(const std::vector<BYTE>& data, UINT fmt) {
    if (fmt == CF_UNICODETEXT && data.size() >= sizeof(wchar_t)) {
        try {
            // Validate data pointer
            if (data.empty() || data.data() == nullptr) {
                return L"[Unicode Text]";
            }
            
            size_t totalChars = data.size() / sizeof(wchar_t);
            if (totalChars == 0 || totalChars > 1000000) { // Limit to 1M chars
                return L"[Unicode Text]";
            }
            
            // Safety check - limit totalChars to prevent excessive processing
            if (totalChars > 10000) {
                return L"[Unicode Text - Large]";
            }
            
            // Find null terminator if present - with bounds checking
            const wchar_t* text = (const wchar_t*)data.data();
            if (!text) return L"[Unicode Text]";
            
            size_t len = totalChars;
            // Limit search to prevent excessive iteration
            size_t searchLimit = std::min(totalChars, (size_t)1000);
            for (size_t i = 0; i < searchLimit && i < totalChars; i++) {
                // Validate we're still within bounds
                if ((i + 1) * sizeof(wchar_t) > data.size()) {
                    len = i;
                    break;
                }
                if (text[i] == L'\0') {
                    len = i;
                    break;
                }
            }
            
            // Limit preview length
            if (len > 50) len = 50;
            
            // Validate that we're not reading past buffer
            if (len > totalChars) len = totalChars;
            if (len * sizeof(wchar_t) > data.size()) {
                len = data.size() / sizeof(wchar_t);
            }
            
            // Create string safely
            if (len > 0 && len <= totalChars) {
                std::wstring preview;
                preview.reserve(len + 3); // Reserve space for "..."
                preview.assign(text, len);
                if (totalChars > 50) preview += L"...";
                return preview;
            }
            return L"[Unicode Text]";
        } catch (const std::bad_alloc&) {
            return L"[Unicode Text]";
        } catch (const std::exception&) {
            return L"[Unicode Text]";
        } catch (...) {
            return L"[Unicode Text]";
        }
    }
    if (fmt == CF_TEXT && data.size() > 0) {
        try {
            size_t len = data.size();
            
            // Find null terminator if present
            const char* text = (const char*)data.data();
            for (size_t i = 0; i < data.size(); i++) {
                if (text[i] == '\0') {
                    len = i;
                    break;
                }
            }
            
            // Limit preview length
            if (len > 50) len = 50;
            
            // Validate that we're not reading past buffer
            if (len > data.size()) len = data.size();
            
            std::string preview(text, len);
            if (data.size() > 50) preview += "...";
            return std::wstring(preview.begin(), preview.end());
        } catch (...) {
            return L"[Text]";
        }
    }
    if (fmt == CF_HDROP) {
        try {
            const std::vector<BYTE>* dropData = GetFormatData(CF_HDROP);
            if (dropData && dropData->size() >= sizeof(DROPFILES)) {
                DROPFILES* df = (DROPFILES*)dropData->data();
                // Validate pFiles offset
                if (df->pFiles > 0 && df->pFiles < (DWORD)dropData->size()) {
                    wchar_t* files = (wchar_t*)((BYTE*)dropData->data() + df->pFiles);
                    size_t maxLen = (dropData->size() - df->pFiles) / sizeof(wchar_t);
                    if (maxLen > 0 && maxLen < 32767) {
                        std::wstring firstFile = files;
                        // Limit to reasonable length
                        if (firstFile.length() > maxLen) {
                            firstFile = firstFile.substr(0, maxLen);
                        }
                        if (firstFile.length() > 30) {
                            firstFile = firstFile.substr(0, 27) + L"...";
                        }
                        return firstFile;
                    }
                }
            }
        } catch (...) {
            // Error processing file drop, return generic
        }
    }
    
    return L"[" + formatName + L"]";
}

// Helper function to convert HBITMAP to DIB data
static std::vector<BYTE> ConvertBitmapToDIB(HBITMAP hBitmap) {
    std::vector<BYTE> dibData;
    if (!hBitmap) return dibData;
    
    HDC hdcScreen = GetDC(nullptr);
    
    BITMAP bm;
    if (GetObject(hBitmap, sizeof(BITMAP), &bm)) {
        BITMAPINFOHEADER bih = {};
        bih.biSize = sizeof(BITMAPINFOHEADER);
        bih.biWidth = bm.bmWidth;
        bih.biHeight = bm.bmHeight;
        bih.biPlanes = 1;
        bih.biBitCount = bm.bmBitsPixel;
        bih.biCompression = BI_RGB;
        
        // Calculate size needed
        int rowSize = ((bm.bmWidth * bm.bmBitsPixel + 31) / 32) * 4;
        int imageSize = rowSize * abs(bm.bmHeight);
        bih.biSizeImage = imageSize;
        
        // Allocate buffer for header + color table + bits
        int colorTableSize = 0;
        if (bih.biBitCount <= 8) {
            colorTableSize = (1 << bih.biBitCount) * sizeof(RGBQUAD);
        }
        
        dibData.resize(sizeof(BITMAPINFOHEADER) + colorTableSize + imageSize);
        BITMAPINFOHEADER* pBih = (BITMAPINFOHEADER*)dibData.data();
        *pBih = bih;
        
        // Get bits
        BITMAPINFO* pBmi = (BITMAPINFO*)dibData.data();
        void* pBits = dibData.data() + sizeof(BITMAPINFOHEADER) + colorTableSize;
        
        if (GetDIBits(hdcScreen, hBitmap, 0, abs(bm.bmHeight), pBits, pBmi, DIB_RGB_COLORS)) {
            // Success
        } else {
            dibData.clear();
        }
    }
    
    ReleaseDC(nullptr, hdcScreen);
    return dibData;
}

HBITMAP ClipboardItem::CreateBitmapFromData() {
    HDC hdcScreen = nullptr;
    HBITMAP hBitmap = nullptr;
    
    try {
        hdcScreen = GetDC(nullptr);
        if (!hdcScreen) return nullptr;
        
        // Try to get DIB data from formats map
        const std::vector<BYTE>* dibData = nullptr;
        UINT dibFormat = 0;
        
        // Try CF_DIBV5 first, then CF_DIB
        dibData = GetFormatData(CF_DIBV5);
        if (dibData && dibData->size() >= sizeof(BITMAPV5HEADER)) {
            dibFormat = CF_DIBV5;
        } else {
            dibData = GetFormatData(CF_DIB);
            if (dibData && dibData->size() >= sizeof(BITMAPINFOHEADER)) {
                dibFormat = CF_DIB;
            }
        }
        
        if (!dibData || dibData->empty() || dibData->data() == nullptr) {
            ReleaseDC(nullptr, hdcScreen);
            return nullptr;
        }
        
        if (dibFormat == CF_DIB || dibFormat == CF_DIBV5) {
            // Validate data size and pointer
            if (dibData->size() >= sizeof(BITMAPINFOHEADER)) {
                const BITMAPINFOHEADER* bmi = (const BITMAPINFOHEADER*)dibData->data();
                
                // Validate header
                if (bmi && bmi->biSize >= sizeof(BITMAPINFOHEADER) && bmi->biSize <= dibData->size()) {
                    // Validate dimensions
                    if (bmi->biWidth > 0 && bmi->biWidth < 50000 && 
                        abs(bmi->biHeight) > 0 && abs(bmi->biHeight) < 50000) {
                        
                        // Calculate bits pointer (accounting for color table)
                        int colorTableSize = 0;
                        if (bmi->biBitCount <= 8 && bmi->biBitCount > 0) {
                            int maxColors = bmi->biClrUsed ? bmi->biClrUsed : (1 << bmi->biBitCount);
                            if (maxColors > 0 && maxColors < 256) {
                                colorTableSize = maxColors * sizeof(RGBQUAD);
                                // Limit color table size
                                if (colorTableSize > 1024 * 1024) colorTableSize = 1024 * 1024;
                            }
                        }
                        
                        // Validate that we have enough data
                        int headerSize = (dibFormat == CF_DIBV5) ? sizeof(BITMAPV5HEADER) : sizeof(BITMAPINFOHEADER);
                        int expectedSize = headerSize + colorTableSize;
                        
                        if (dibData->size() >= (size_t)expectedSize && expectedSize > 0) {
                            const void* bits = (const BYTE*)dibData->data() + expectedSize;
                            if (bits && (BYTE*)bits < dibData->data() + dibData->size()) {
                                hBitmap = CreateDIBitmap(hdcScreen, bmi, CBM_INIT, (void*)bits, (BITMAPINFO*)bmi, DIB_RGB_COLORS);
                            }
                        }
                    }
                }
            }
        }
        
        ReleaseDC(nullptr, hdcScreen);
    } catch (...) {
        if (hdcScreen) ReleaseDC(nullptr, hdcScreen);
        return nullptr;
    }
    
    return hBitmap;
}

void ClipboardItem::GenerateThumbnail() {
    // CRITICAL: This function must NEVER crash - wrap everything in try-catch
    try {
        if (format == CF_BITMAP || format == CF_DIB || format == CF_DIBV5) {
            // Generate thumbnail for bitmap/DIB
            HDC hdcScreen = nullptr;
            HBITMAP hBitmap = nullptr;
            HDC hdcThumb = nullptr;
            HBITMAP hThumb = nullptr;
            HDC hdcSrc = nullptr;
            HGDIOBJ oldThumb = nullptr;
            HGDIOBJ oldSrc = nullptr;
            
            try {
                hdcScreen = GetDC(nullptr);
                if (!hdcScreen) return;
                
                hBitmap = CreateBitmapFromData();
                if (!hBitmap) {
                    ReleaseDC(nullptr, hdcScreen);
                    return;
                }
                
                // Create thumbnail (48x48)
                const int thumbSize = 48;
                hdcThumb = CreateCompatibleDC(hdcScreen);
                if (!hdcThumb) {
                    DeleteObject(hBitmap);
                    ReleaseDC(nullptr, hdcScreen);
                    return;
                }
                
                hThumb = CreateCompatibleBitmap(hdcScreen, thumbSize, thumbSize);
                if (!hThumb) {
                    DeleteDC(hdcThumb);
                    DeleteObject(hBitmap);
                    ReleaseDC(nullptr, hdcScreen);
                    return;
                }
                
                oldThumb = SelectObject(hdcThumb, hThumb);
                
                // Get bitmap dimensions
                BITMAP bm = {0};
                if (GetObject(hBitmap, sizeof(BITMAP), &bm) && bm.bmWidth > 0 && bm.bmHeight > 0) {
                    // Validate dimensions
                    int srcW = bm.bmWidth;
                    int srcH = bm.bmHeight;
                    if (srcW > 0 && srcH > 0 && srcW < 10000 && srcH < 10000) {
                        float scale = std::min((float)thumbSize / srcW, (float)thumbSize / srcH);
                        int dstW = (int)(srcW * scale);
                        int dstH = (int)(srcH * scale);
                        int offsetX = (thumbSize - dstW) / 2;
                        int offsetY = (thumbSize - dstH) / 2;
                        
                        // Fill background - Black OLED theme
                        RECT bgRect = {0, 0, thumbSize, thumbSize};
                        FillRect(hdcThumb, &bgRect, (HBRUSH)GetStockObject(BLACK_BRUSH));
                        
                        // Stretch bitmap to thumbnail
                        hdcSrc = CreateCompatibleDC(hdcScreen);
                        if (hdcSrc) {
                            oldSrc = SelectObject(hdcSrc, hBitmap);
                            SetStretchBltMode(hdcThumb, HALFTONE);
                            StretchBlt(hdcThumb, offsetX, offsetY, dstW, dstH, hdcSrc, 0, 0, srcW, srcH, SRCCOPY);
                            SelectObject(hdcSrc, oldSrc);
                            DeleteDC(hdcSrc);
                        }
                        
                        SelectObject(hdcThumb, oldThumb);
                        thumbnail = hThumb;
                        hThumb = nullptr; // Don't delete, we're using it
                        
                        // Generate larger preview bitmap for hover preview (max 500x500 for quality)
                        const int previewMaxSize = 500;
                        HDC hdcPreview = CreateCompatibleDC(hdcScreen);
                        if (hdcPreview) {
                            float previewScale = std::min((float)previewMaxSize / srcW, (float)previewMaxSize / srcH);
                            int previewW = (int)(srcW * previewScale);
                            int previewH = (int)(srcH * previewScale);
                            
                            // Limit to reasonable size to prevent memory issues
                            if (previewW > 0 && previewH > 0 && previewW <= 1000 && previewH <= 1000) {
                                HBITMAP hPreview = CreateCompatibleBitmap(hdcScreen, previewW, previewH);
                                if (hPreview) {
                                    HGDIOBJ oldPreview = SelectObject(hdcPreview, hPreview);
                                    
                                    // Fill background - Black OLED theme
                                    RECT bgRect = {0, 0, previewW, previewH};
                                    FillRect(hdcPreview, &bgRect, (HBRUSH)GetStockObject(BLACK_BRUSH));
                                    
                                    // Stretch original bitmap to preview size with high quality
                                    hdcSrc = CreateCompatibleDC(hdcScreen);
                                    if (hdcSrc) {
                                        oldSrc = SelectObject(hdcSrc, hBitmap);
                                        SetStretchBltMode(hdcPreview, HALFTONE);
                                        StretchBlt(hdcPreview, 0, 0, previewW, previewH, hdcSrc, 0, 0, srcW, srcH, SRCCOPY);
                                        SelectObject(hdcSrc, oldSrc);
                                        DeleteDC(hdcSrc);
                                    }
                                    
                                    SelectObject(hdcPreview, oldPreview);
                                    
                                    // Delete old preview bitmap if it exists
                                    if (previewBitmap) {
                                        DeleteObject(previewBitmap);
                                    }
                                    previewBitmap = hPreview;
                                }
                            }
                            DeleteDC(hdcPreview);
                        }
                    } else {
                        SelectObject(hdcThumb, oldThumb);
                        DeleteObject(hThumb);
                    }
                } else {
                    SelectObject(hdcThumb, oldThumb);
                    DeleteObject(hThumb);
                }
                
                DeleteDC(hdcThumb);
                DeleteObject(hBitmap);
                ReleaseDC(nullptr, hdcScreen);
            } catch (...) {
                // Clean up on exception
                if (oldThumb && hdcThumb) SelectObject(hdcThumb, oldThumb);
                if (oldSrc && hdcSrc) SelectObject(hdcSrc, oldSrc);
                if (hThumb) DeleteObject(hThumb);
                if (hdcThumb) DeleteDC(hdcThumb);
                if (hBitmap) DeleteObject(hBitmap);
                if (hdcSrc) DeleteDC(hdcSrc);
                if (hdcScreen) ReleaseDC(nullptr, hdcScreen);
            }
        } else if (format == CF_HDROP && (isImage || isVideo)) {
            // For image/video files, extract thumbnail using Windows Shell API
            const std::vector<BYTE>* dropData = GetFormatData(CF_HDROP);
            if (dropData && dropData->size() >= sizeof(DROPFILES)) {
                try {
                    DROPFILES* df = (DROPFILES*)dropData->data();
                    // Validate pFiles offset
                    if (df->pFiles > 0 && df->pFiles < (DWORD)dropData->size()) {
                        wchar_t* files = (wchar_t*)((BYTE*)dropData->data() + df->pFiles);
                        std::wstring firstFile = files;
                        
                        // Use Windows Shell to extract thumbnail
                        IShellItemImageFactory* pImageFactory = nullptr;
                        HRESULT hr = SHCreateItemFromParsingName(firstFile.c_str(), nullptr, IID_PPV_ARGS(&pImageFactory));
                        if (SUCCEEDED(hr) && pImageFactory) {
                            HBITMAP hBitmap = nullptr;
                            SIZE size = {48, 48};
                            hr = pImageFactory->GetImage(size, SIIGBF_THUMBNAILONLY, &hBitmap);
                            if (SUCCEEDED(hr) && hBitmap) {
                                thumbnail = hBitmap;
                            }
                            pImageFactory->Release();
                            pImageFactory = nullptr;
                        }
                    }
                } catch (...) {
                    // Error generating thumbnail, continue without thumbnail
                }
            }
        }
    } catch (...) {
        // Error generating thumbnail, continue without thumbnail
    }
}

HBITMAP ClipboardItem::GetPreviewBitmap() {
    // Return larger preview bitmap if available, otherwise fall back to thumbnail
    if (previewBitmap) {
        return previewBitmap;
    }
    return thumbnail;
}

std::wstring ClipboardItem::GetFullSearchableText() const {
    const size_t maxSearchBytes = 500000;  // Limit for performance
    const std::vector<BYTE>* udata = GetFormatData(CF_UNICODETEXT);
    if (udata && !udata->empty() && udata->size() <= maxSearchBytes) {
        size_t len = udata->size() / sizeof(wchar_t);
        if (len > 0) {
            const wchar_t* p = (const wchar_t*)udata->data();
            return std::wstring(p, len);
        }
    }
    const std::vector<BYTE>* adata = GetFormatData(CF_TEXT);
    if (adata && !adata->empty() && adata->size() <= maxSearchBytes) {
        int wlen = MultiByteToWideChar(CP_ACP, 0, (const char*)adata->data(), (int)adata->size(), nullptr, 0);
        if (wlen > 0) {
            std::wstring result(wlen, 0);
            MultiByteToWideChar(CP_ACP, 0, (const char*)adata->data(), (int)adata->size(), &result[0], wlen);
            return result;
        }
    }
    const std::vector<BYTE>* oemdata = GetFormatData(CF_OEMTEXT);
    if (oemdata && !oemdata->empty() && oemdata->size() <= maxSearchBytes) {
        int wlen = MultiByteToWideChar(CP_OEMCP, 0, (const char*)oemdata->data(), (int)oemdata->size(), nullptr, 0);
        if (wlen > 0) {
            std::wstring result(wlen, 0);
            MultiByteToWideChar(CP_OEMCP, 0, (const char*)oemdata->data(), (int)oemdata->size(), &result[0], wlen);
            return result;
        }
    }
    return L"";
}

// Get normalized text from an item for duplicate detection (content up to first null, trimmed).
// So "text" copied twice with different trailing nulls still counts as duplicate.
static std::wstring GetNormalizedTextForDuplicateCheck(const ClipboardItem* item) {
    if (!item) return L"";
    const std::vector<BYTE>* udata = item->GetFormatData(CF_UNICODETEXT);
    if (udata && udata->size() >= sizeof(wchar_t)) {
        size_t len = udata->size() / sizeof(wchar_t);
        const wchar_t* p = (const wchar_t*)udata->data();
        size_t n = 0;
        while (n < len && p[n] != L'\0') n++;
        std::wstring s(p, n);
        while (!s.empty() && (s.back() == L' ' || s.back() == L'\t' || s.back() == L'\r' || s.back() == L'\n'))
            s.pop_back();
        return s;
    }
    const std::vector<BYTE>* adata = item->GetFormatData(CF_TEXT);
    if (adata && !adata->empty()) {
        int wlen = MultiByteToWideChar(CP_ACP, 0, (const char*)adata->data(), (int)adata->size(), nullptr, 0);
        if (wlen > 0) {
            std::wstring result(wlen, 0);
            MultiByteToWideChar(CP_ACP, 0, (const char*)adata->data(), (int)adata->size(), &result[0], wlen);
            size_t n = 0;
            while (n < result.size() && result[n] != L'\0') n++;
            result.resize(n);
            while (!result.empty() && (result.back() == L' ' || result.back() == L'\t' || result.back() == L'\r' || result.back() == L'\n'))
                result.pop_back();
            return result;
        }
    }
    return L"";
}

// Plain text from a history item for direct injection (paste without system clipboard).
// Uses full UTF-16 length from stored bytes (trim trailing nulls only) so content matches UIA / clipboard.
static std::wstring GetPlainTextForDirectPaste(const ClipboardItem* item) {
    if (!item) return L"";
    const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
    if (textData && textData->size() >= sizeof(wchar_t)) {
        size_t wcharCount = textData->size() / sizeof(wchar_t);
        const wchar_t* p = (const wchar_t*)textData->data();
        while (wcharCount > 0 && p[wcharCount - 1] == L'\0')
            wcharCount--;
        return std::wstring(p, wcharCount);
    }
    textData = item->GetFormatData(CF_TEXT);
    if (textData && !textData->empty()) {
        int wlen = MultiByteToWideChar(CP_ACP, 0, (const char*)textData->data(), (int)textData->size(), nullptr, 0);
        if (wlen > 0) {
            std::wstring result(wlen, 0);
            MultiByteToWideChar(CP_ACP, 0, (const char*)textData->data(), (int)textData->size(), &result[0], wlen);
            size_t n = 0;
            while (n < result.size() && result[n] != L'\0') n++;
            result.resize(n);
            return result;
        }
    }
    if (!item->preview.empty() && item->preview[0] != L'[')
        return item->preview;
    return L"";
}

static void SendKeyUpIfDown(WORD vk) {
    if (!(GetAsyncKeyState(vk) & 0x8000)) return;
    INPUT in = {};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = vk;
    in.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &in, sizeof(INPUT));
}

// Release keys still held from Ctrl+F11 / Ctrl+Shift+F11 so typing or Ctrl+V is not corrupted.
// Use explicit L/R virtual keys — generic VK_CONTROL / VK_SHIFT keyup can miss the physical key still down.
static void ReleaseHotkeyModifiersForPaste() {
    SendKeyUpIfDown(VK_F11);
    SendKeyUpIfDown(VK_LSHIFT);
    SendKeyUpIfDown(VK_RSHIFT);
    SendKeyUpIfDown(VK_LCONTROL);
    SendKeyUpIfDown(VK_RCONTROL);
    Sleep(20);
}

// Release every modifier that can change how injected Unicode is interpreted (Ctrl/Alt/Shift/Win).
static void ReleaseAllModifierKeysForKeystrokePaste() {
    SendKeyUpIfDown(VK_MENU);  // generic Alt
    SendKeyUpIfDown(VK_LMENU);
    SendKeyUpIfDown(VK_RMENU);
    SendKeyUpIfDown(VK_CONTROL);
    SendKeyUpIfDown(VK_LCONTROL);
    SendKeyUpIfDown(VK_RCONTROL);
    SendKeyUpIfDown(VK_SHIFT);
    SendKeyUpIfDown(VK_LSHIFT);
    SendKeyUpIfDown(VK_RSHIFT);
    SendKeyUpIfDown(VK_LWIN);
    SendKeyUpIfDown(VK_RWIN);
    Sleep(25);
}

static bool SetClipboardUnicodeOnly(HWND hwnd, const std::wstring& text) {
    if (!OpenClipboard(hwnd)) return false;
    EmptyClipboard();
    size_t byteLen = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, byteLen);
    if (!hMem) { CloseClipboard(); return false; }
    void* p = GlobalLock(hMem);
    if (!p) { GlobalFree(hMem); CloseClipboard(); return false; }
    memcpy(p, text.c_str(), byteLen);
    GlobalUnlock(hMem);
    if (!SetClipboardData(CF_UNICODETEXT, hMem)) { GlobalFree(hMem); CloseClipboard(); return false; }
    CloseClipboard();
    return true;
}

// Full clipboard backup for swap+paste (not just CF_UNICODETEXT). CF_BITMAP is stored as CF_DIB bytes.
static bool BackupClipboardSerialFormats(HWND hwnd, std::map<UINT, std::vector<BYTE>>& out) {
    out.clear();
    if (!OpenClipboard(hwnd)) return false;
    const size_t MAX_TOTAL = 50 * 1024 * 1024;
    size_t total = 0;
    auto skipUnsupported = [](UINT f) {
        return f == CF_PALETTE || f == CF_METAFILEPICT || f == CF_ENHMETAFILE ||
               f == 0x0082 || f == 0x008E || f == 0x0083;
    };
    UINT fmt = 0;
    while ((fmt = EnumClipboardFormats(fmt)) != 0) {
        if (skipUnsupported(fmt)) continue;
        if (fmt == CF_BITMAP) {
            if (out.find(CF_DIB) != out.end()) continue;
            HBITMAP hBmp = (HBITMAP)GetClipboardData(CF_BITMAP);
            if (!hBmp) continue;
            std::vector<BYTE> data = ConvertBitmapToDIB(hBmp);
            if (data.empty() || total + data.size() > MAX_TOTAL) continue;
            out[CF_DIB] = std::move(data);
            total += out[CF_DIB].size();
            continue;
        }
        HGLOBAL hMem = GetClipboardData(fmt);
        if (!hMem) continue;
        SIZE_T sz = GlobalSize(hMem);
        if (sz == 0 || sz >= 100 * 1024 * 1024) continue;
        void* p = GlobalLock(hMem);
        if (!p) continue;
        std::vector<BYTE> data(sz);
        memcpy(data.data(), p, sz);
        GlobalUnlock(hMem);
        if (total + data.size() > MAX_TOTAL) continue;
        if (out.find(fmt) != out.end()) continue;
        out[fmt] = std::move(data);
        total += out[fmt].size();
    }
    CloseClipboard();
    return true;
}

static bool RestoreClipboardSerialFormats(HWND hwnd, const std::map<UINT, std::vector<BYTE>>& backup) {
    if (!OpenClipboard(hwnd)) return false;
    EmptyClipboard();
    for (const auto& kv : backup) {
        if (kv.second.empty()) continue;
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, kv.second.size());
        if (!hMem) continue;
        void* p = GlobalLock(hMem);
        if (!p) {
            GlobalFree(hMem);
            continue;
        }
        memcpy(p, kv.second.data(), kv.second.size());
        GlobalUnlock(hMem);
        if (!SetClipboardData(kv.first, hMem))
            GlobalFree(hMem);
    }
    CloseClipboard();
    return true;
}

// Restore all formats from a history item (same strategy as PasteItem rich paste).
static bool SetClipboardFromHistoryItem(HWND hwnd, const ClipboardItem* item) {
    if (!item || item->formats.empty()) return false;
    if (!OpenClipboard(hwnd)) return false;
    EmptyClipboard();
    bool success = false;
    for (const auto& formatPair : item->formats) {
        UINT fmt = formatPair.first;
        const std::vector<BYTE>& formatData = formatPair.second;
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, formatData.size());
        if (!hMem) continue;
        void* pMem = GlobalLock(hMem);
        if (!pMem) {
            GlobalFree(hMem);
            continue;
        }
        memcpy(pMem, formatData.data(), formatData.size());
        GlobalUnlock(hMem);
        if (SetClipboardData(fmt, hMem))
            success = true;
        else
            GlobalFree(hMem);
    }
    CloseClipboard();
    return success;
}

static std::wstring LastPastedTextFromItem(const ClipboardItem* item) {
    if (!item) return L"";
    const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
    if (textData && textData->size() >= sizeof(wchar_t)) {
        size_t len = textData->size() / sizeof(wchar_t);
        const wchar_t* textPtr = (const wchar_t*)textData->data();
        size_t actualLen = len;
        for (size_t j = 0; j < len; j++) {
            if (textPtr[j] == L'\0') {
                actualLen = j;
                break;
            }
        }
        if (actualLen > 0) return std::wstring(textPtr, actualLen);
    }
    return item->preview;
}

// One SendInput batch so Ctrl is down before V; keybd_event can interleave badly with our LL hook.
static void SendCtrlV() {
    INPUT in[4] = {};
    in[0].type = INPUT_KEYBOARD;
    in[0].ki.wVk = VK_LCONTROL;
    in[1].type = INPUT_KEYBOARD;
    in[1].ki.wVk = 'V';
    in[2].type = INPUT_KEYBOARD;
    in[2].ki.wVk = 'V';
    in[2].ki.dwFlags = KEYEVENTF_KEYUP;
    in[3].type = INPUT_KEYBOARD;
    in[3].ki.wVk = VK_LCONTROL;
    in[3].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(4, in, sizeof(INPUT));
}

// Cached registered clipboard formats for rich text/HTML.
static UINT CfRtf() {
    static UINT cf = 0;
    if (cf == 0) cf = RegisterClipboardFormatA("Rich Text Format");
    return cf;
}
static UINT CfHtml() {
    static UINT cf = 0;
    if (cf == 0) cf = RegisterClipboardFormatA("HTML Format");
    return cf;
}

// Robust OpenClipboard with retries. Some apps hold the clipboard briefly when copying.
static bool OpenClipboardWithRetry(HWND hwnd, int attempts = 24, int sleepMs = 8) {
    for (int i = 0; i < attempts; i++) {
        if (OpenClipboard(hwnd)) return true;
        Sleep(sleepMs);
    }
    return false;
}

// Allocate moveable global memory, copy bytes, and put it on the clipboard for `fmt`.
// Clipboard must be open. Returns true on success (clipboard takes ownership).
static bool PutBytesOnClipboard(UINT fmt, const void* data, size_t size) {
    if (!data || size == 0) return false;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
    if (!hMem) return false;
    void* p = GlobalLock(hMem);
    if (!p) { GlobalFree(hMem); return false; }
    memcpy(p, data, size);
    GlobalUnlock(hMem);
    if (!SetClipboardData(fmt, hMem)) {
        GlobalFree(hMem);
        return false;
    }
    return true;
}

// Convert a "stored RTF" wstring (where each wchar holds one RTF byte 0..255) back into
// a byte string suitable for the "Rich Text Format" clipboard format.
static std::string RtfBytesFromStoredWide(const std::wstring& packed) {
    std::string out;
    out.reserve(packed.size());
    for (wchar_t wc : packed) out += (char)(wc & 0xFF);
    return out;
}

// True when an item carries text/RTF (anything we can combine into a single paste payload).
static bool ItemHasOnlyTextFormats(const ClipboardItem* item) {
    if (!item) return false;
    UINT rtf = CfRtf();
    UINT html = CfHtml();
    for (const auto& kv : item->formats) {
        UINT f = kv.first;
        if (f == CF_UNICODETEXT || f == CF_TEXT || f == CF_OEMTEXT || f == CF_LOCALE) continue;
        if (rtf && f == rtf) continue;
        if (html && f == html) continue;
        return false;
    }
    return true;
}

// Build a minimal RTF document wrapping `text` as a single paragraph block. Used to seed
// rich-text-aware targets when concatenating multiple plain items.
static std::string BuildRtfWrapFromUnicode(const std::wstring& text) {
    std::string s;
    s.reserve(text.size() * 3 + 128);
    s += "{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033"
         "{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}}"
         "\\f0\\fs22 ";
    for (size_t i = 0; i < text.size(); i++) {
        wchar_t wc = text[i];
        if (wc == L'\r') continue;
        if (wc == L'\n') { s += "\\par\n"; continue; }
        if (wc == L'\\' || wc == L'{' || wc == L'}') { s += '\\'; s += (char)wc; continue; }
        if (wc < 0x80 && wc >= 0x20) { s += (char)wc; continue; }
        if (wc == L'\t') { s += "\\tab "; continue; }
        char buf[24];
        int code = (int)(short)wc;  // RTF uses signed 16-bit
        snprintf(buf, sizeof(buf), "\\u%d? ", code);
        s += buf;
    }
    s += "}";
    return s;
}

// Type plain text into the focused control without using the clipboard (Unicode keyboard events).
// Improvements over the previous version:
//   - Small batches (24 inputs) so the receiving app's input queue does not drop events.
//   - Partial SendInput results are retried from the next index instead of dropping the tail.
//   - Brief pacing between batches; \r/\n converted to a real VK_RETURN keystroke so rich
//     editors (Word, OneNote, browsers) get real Enter events instead of literal U+000A.
static bool SendUnicodeTextAsKeystrokes(const std::wstring& text) {
    ReleaseAllModifierKeysForKeystrokePaste();
    if (text.empty()) return true;

    auto pushUnicode = [](std::vector<INPUT>& out, wchar_t ch) {
        INPUT down{};
        down.type = INPUT_KEYBOARD;
        down.ki.wVk = 0;
        down.ki.wScan = ch;
        down.ki.dwFlags = KEYEVENTF_UNICODE;
        INPUT up = down;
        up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        out.push_back(down);
        out.push_back(up);
    };
    auto pushVK = [](std::vector<INPUT>& out, WORD vk) {
        INPUT down{};
        down.type = INPUT_KEYBOARD;
        down.ki.wVk = vk;
        INPUT up = down;
        up.ki.dwFlags = KEYEVENTF_KEYUP;
        out.push_back(down);
        out.push_back(up);
    };

    auto flushBatch = [](std::vector<INPUT>& b) -> bool {
        if (b.empty()) return true;
        UINT total = (UINT)b.size();
        UINT idx = 0;
        for (int attempt = 0; attempt < 8 && idx < total; attempt++) {
            UINT sent = SendInput(total - idx, b.data() + idx, sizeof(INPUT));
            if (sent == 0) {
                DWORD err = GetLastError();
                (void)err;
                Sleep(10);
                continue;
            }
            idx += sent;
            if (idx < total) Sleep(4);
        }
        b.clear();
        return idx == total;
    };

    const size_t kBatchCap = 24;  // 12 chars per batch max
    std::vector<INPUT> batch;
    batch.reserve(kBatchCap + 4);

    size_t i = 0;
    while (i < text.size()) {
        wchar_t ch = text[i];

        // \r\n → single VK_RETURN; lone \n/\r → VK_RETURN; literal \t → VK_TAB.
        if (ch == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n') {
            if (batch.size() + 2 > kBatchCap) { if (!flushBatch(batch)) return false; Sleep(2); }
            pushVK(batch, VK_RETURN);
            i += 2;
            continue;
        }
        if (ch == L'\r' || ch == L'\n') {
            if (batch.size() + 2 > kBatchCap) { if (!flushBatch(batch)) return false; Sleep(2); }
            pushVK(batch, VK_RETURN);
            i += 1;
            continue;
        }
        if (ch == L'\t') {
            if (batch.size() + 2 > kBatchCap) { if (!flushBatch(batch)) return false; Sleep(2); }
            pushVK(batch, VK_TAB);
            i += 1;
            continue;
        }

        // Surrogate pair: 4 inputs must go together in one batch.
        bool isHighSurrogate = (ch >= 0xD800 && ch <= 0xDBFF);
        bool pair = isHighSurrogate && (i + 1) < text.size() &&
            text[i + 1] >= 0xDC00 && text[i + 1] <= 0xDFFF;
        if (pair) {
            if (batch.size() + 4 > kBatchCap) { if (!flushBatch(batch)) return false; Sleep(2); }
            pushUnicode(batch, ch);
            pushUnicode(batch, text[i + 1]);
            i += 2;
        } else {
            if (ch == L'\0') { i += 1; continue; }
            if (batch.size() + 2 > kBatchCap) { if (!flushBatch(batch)) return false; Sleep(2); }
            pushUnicode(batch, ch);
            i += 1;
        }

        // Larger texts: yield briefly every ~256 chars to keep the foreground app responsive.
        if ((i & 0xFF) == 0) Sleep(1);
    }
    return flushBatch(batch);
}

// ClipboardManager implementation
ClipboardManager::ClipboardManager()
    : hwndMain(nullptr), hwndList(nullptr), hwndPreview(nullptr), hwndSearch(nullptr), hwndSettings(nullptr), hwndSnippetsManager(nullptr), isRunning(false), listVisible(false), lastSequenceNumber(0), hKeyboardHook(nullptr), scrollOffset(0), itemsPerPage(10), numberInput(L""), searchText(L""), snippetsMode(false), lastSKeyTime(0), ignoreNextSChar(false), isPasting(false), isProcessingClipboard(false), lastPastedText(L""), previousFocusWindow(nullptr), hoveredItemIndex(-1), selectedIndex(0), multiSelectAnchor(-1), originalSearchEditProc(nullptr), lastHotkeyTick(0), hasImmediateClipboardSnapshot(false), maxItems(DEFAULT_MAX_ITEMS) {
    instance = this;
    ZeroMemory(&nid, sizeof(nid));
    hotkeyConfig.modifiers = MOD_CONTROL;
    hotkeyConfig.vkCode = VK_DECIMAL;
    // Snippets overlay has no default shortcut. The overlay can still be opened from the
    // tray menu, from the clipboard overlay via Ctrl+Right, or by typing *set + Enter.
    // Users can assign a custom binding via Settings if they want one.
    snippetsHotkey.modifiers = 0;
    snippetsHotkey.vkCode = 0;
    copyFocusedHotkey.modifiers = MOD_CONTROL;
    copyFocusedHotkey.vkCode = VK_F10;
    pasteFocusedHotkey.modifiers = MOD_CONTROL;
    pasteFocusedHotkey.vkCode = VK_F11;
    pasteClipboardHotkey.modifiers = MOD_CONTROL | MOD_SHIFT;
    pasteClipboardHotkey.vkCode = VK_F11;
}

ClipboardManager::~ClipboardManager() {
    Stop();
}

bool ClipboardManager::Initialize() {
    // Initialize COM for Shell API (for thumbnails)
    // Use COINIT_MULTITHREADED for better compatibility, or check if already initialized
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    // Ignore RPC_E_CHANGED_MODE error (COM already initialized with different mode)
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
        // COM initialization failed, but continue anyway
    }

    // Apply the saved theme BEFORE registering window classes — class background brushes
    // capture the color at registration time. Font color/face overrides are loaded here
    // so the very first overlay paint uses the chosen font and accent color.
    g_themeFontColor = LoadThemeFontColorFromRegistry();
    std::wstring savedFace = LoadThemeFontFaceFromRegistry();
    if (!savedFace.empty() && FontFamilyExists(savedFace))
        g_themeFontFace = savedFace;
    else
        g_themeFontFace = kDefaultThemeFontFace;
    ResetOverlayFontCache();
    LoadThemeColorsFromRegistry();
    ApplyThemeId(LoadThemeIdFromRegistry());
    maxItems = LoadMaxItemsFromRegistry(DEFAULT_MAX_ITEMS, MIN_MAX_ITEMS, MAX_MAX_ITEMS);

    // Register window class for main window
    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = GetModuleHandle(nullptr);
    wc.lpszClassName = L"ClipboardManagerClass";
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    
    if (!RegisterClass(&wc)) {
        return false;
    }
    
    // Register window class for list window
    WNDCLASS wcList = {};
    wcList.lpfnWndProc = ListWindowProc;
    wcList.hInstance = GetModuleHandle(nullptr);
    wcList.lpszClassName = L"ClipboardManagerListClass";
    wcList.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wcList.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcList.hbrBackground = CreateSolidBrush(Theme5250::BG);
    wcList.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    
    if (!RegisterClass(&wcList)) {
        return false;
    }
    
    // Register window class for preview window
    WNDCLASS wcPreview = {};
    wcPreview.lpfnWndProc = ListWindowProc; // Reuse same proc, we'll handle it
    wcPreview.hInstance = GetModuleHandle(nullptr);
    wcPreview.lpszClassName = L"ClipboardManagerPreviewClass";
    wcPreview.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    wcPreview.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcPreview.hbrBackground = CreateSolidBrush(Theme5250::BG);
    wcPreview.style = CS_HREDRAW | CS_VREDRAW;
    
    RegisterClass(&wcPreview);
    
    // Create main window (hidden but must exist for hotkeys)
    hwndMain = CreateWindowEx(
        0,
        L"ClipboardManagerClass",
        L"clip2",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 100, 100,
        nullptr, nullptr, GetModuleHandle(nullptr), nullptr
    );
    
    if (!hwndMain) {
        return false;
    }
    
    // Hide the window but keep it available for messages
    ShowWindow(hwndMain, SW_HIDE);
    
    // Create list window (initially hidden) - no scrollbar (scrolling via keys/mouse wheel)
    hwndList = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        L"ClipboardManagerListClass",
        L"clip2",
        WS_POPUP, // Removed WS_BORDER and WS_VSCROLL - we'll draw rounded border ourselves
        CW_USEDEFAULT, CW_USEDEFAULT,
        WINDOW_WIDTH, WINDOW_HEIGHT,
        nullptr, nullptr, GetModuleHandle(nullptr), nullptr
    );
    
    if (!hwndList) {
        return false;
    }
    
    // Set window region - sharp corners for 5250 terminal look
    HRGN hRgn = CreateRectRgn(0, 0, WINDOW_WIDTH + 1, WINDOW_HEIGHT + 1);
    SetWindowRgn(hwndList, hRgn, TRUE);
    DeleteObject(hRgn);
    
    ShowWindow(hwndList, SW_HIDE);
    
    // Create search edit control (child of list window) - Dark theme style
    hwndSearch = CreateWindowEx(
        0, // Remove WS_EX_CLIENTEDGE for modern look, we'll draw border ourselves
        L"EDIT",
        L"",
        WS_CHILD | WS_VISIBLE | ES_LEFT | ES_AUTOHSCROLL,
        5, 30, WINDOW_WIDTH - 10, 25,
        hwndList,
        nullptr,
        GetModuleHandle(nullptr),
        nullptr
    );
    
    if (hwndSearch) {
        SendMessage(hwndSearch, WM_SETFONT, (WPARAM)GetOverlayFont(), TRUE);
        SendMessage(hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search... (Ctrl+F)");
        // Subclass the edit control to handle Enter key
        originalSearchEditProc = (WNDPROC)SetWindowLongPtr(hwndSearch, GWLP_WNDPROC, (LONG_PTR)SearchEditProc);
        ShowWindow(hwndSearch, SW_HIDE);  // Hide initially, show when list is visible
    }
    
    // Create preview window (initially hidden) with rounded borders
    // Larger size (400x400) to better show image previews
    const int PREVIEW_WINDOW_SIZE = 400;
    hwndPreview = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED,
        L"ClipboardManagerPreviewClass",
        L"Preview",
        WS_POPUP, // Removed WS_BORDER - we'll draw rounded border ourselves
        CW_USEDEFAULT, CW_USEDEFAULT,
        PREVIEW_WINDOW_SIZE, PREVIEW_WINDOW_SIZE,
        nullptr, nullptr, GetModuleHandle(nullptr), nullptr
    );
    
    if (hwndPreview) {
        // Set window region for preview - sharp corners for 5250 terminal look
        HRGN hPreviewRgn = CreateRectRgn(0, 0, PREVIEW_WINDOW_SIZE + 1, PREVIEW_WINDOW_SIZE + 1);
        SetWindowRgn(hwndPreview, hPreviewRgn, TRUE);
        DeleteObject(hPreviewRgn);
        
        ShowWindow(hwndPreview, SW_HIDE);
        SetLayeredWindowAttributes(hwndPreview, 0, 255, LWA_ALPHA);
    }
    
    CreateTrayIcon();
    
    // Load hotkey configuration from registry
    LoadHotkeyConfig();
    
    // Load snippets from registry
    LoadSnippets();
    
    // Load persisted clipboard history
    LoadClipboardHistory();
    
    InstallKeyboardHook();
    RegisterHotkey();  // Global hotkey as backup when hook doesn't fire (e.g. some fullscreen/admin apps)
    
    wmTaskbarCreated = RegisterWindowMessage(L"TaskbarCreated");
    lastSequenceNumber = GetClipboardSequenceNumber();
    
    // Register for clipboard update notifications (instant notification when clipboard changes)
    AddClipboardFormatListener(hwndMain);
    
    // Warm up MCI so first copy plays sound (cold MCI often fails on first use)
    WarmUpClickSound();
    
    return true;
}

void ClipboardManager::Run() {
    isRunning = true;
    MSG msg;
    
    while (isRunning) {
        // Process all pending messages first (important for hotkeys)
        // Use GetMessage for better hotkey message delivery
        BOOL bRet = GetMessage(&msg, nullptr, 0, 0);
        if (bRet == -1) {
            // Error occurred
            break;
        } else if (bRet == 0) {
            // WM_QUIT received
            isRunning = false;
            break;
        }
        
        TranslateMessage(&msg);
        DispatchMessage(&msg);
        
        // Process any remaining messages in queue
        while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                isRunning = false;
                return;
            }
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        
        // Sleep briefly to avoid busy-waiting
        Sleep(10);
    }
}

void ClipboardManager::Stop() {
    isRunning = false;
    SaveClipboardHistory();
    HideListWindow();
    RemoveTrayIcon();
    UninstallKeyboardHook();
    UnregisterHotkey();
    
    // Unregister clipboard listener
    if (hwndMain) {
        RemoveClipboardFormatListener(hwndMain);
    }
    
    // Clean shutdown
    if (hwndPreview) {
        DestroyWindow(hwndPreview);
        hwndPreview = nullptr;
    }
    if (hwndList) {
        DestroyWindow(hwndList);
        hwndList = nullptr;
    }
    if (hwndMain) {
        DestroyWindow(hwndMain);
        hwndMain = nullptr;
    }
    
    // Uninitialize COM
    CoUninitialize();
}

LRESULT CALLBACK ClipboardManager::WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    
    if (!mgr) {
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    
    if (uMsg == mgr->wmTaskbarCreated && mgr->wmTaskbarCreated) {
        mgr->CreateTrayIcon();
        return 0;
    }
    
    switch (uMsg) {
    case WM_TRAYICON:
        if (LOWORD(lParam) == WM_RBUTTONUP) {
            mgr->previousFocusWindow = GetForegroundWindow();  // Save for snippet paste target
            POINT pt;
            GetCursorPos(&pt);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, 1, L"Show Clipboard");
#ifdef HAVE_UIAUTOMATION
            AppendMenu(hMenu, MF_STRING, 6, L"Copy from focused control");
#endif
            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            // Snippets submenu
            if (!mgr->snippets.empty()) {
                HMENU hSnippetsMenu = CreatePopupMenu();
                for (size_t i = 0; i < mgr->snippets.size(); i++) {
                    AppendMenuW(hSnippetsMenu, MF_STRING, (UINT_PTR)(100 + i), mgr->snippets[i].name.c_str());
                }
                AppendMenu(hSnippetsMenu, MF_SEPARATOR, 0, nullptr);
                AppendMenu(hSnippetsMenu, MF_STRING, 5, L"Manage Snippets...");
                AppendMenu(hMenu, MF_STRING | MF_POPUP, (UINT_PTR)hSnippetsMenu, L"Snippets");
            } else {
                AppendMenu(hMenu, MF_STRING, 5, L"Snippets...");
            }
            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            // Add "Start with Windows" option with checkmark
            bool startupEnabled = mgr->IsStartupWithWindows();
            AppendMenu(hMenu, startupEnabled ? (MF_STRING | MF_CHECKED) : MF_STRING, 3, L"Start with Windows");
            AppendMenu(hMenu, MF_STRING, 4, L"Settings");
            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            AppendMenu(hMenu, MF_STRING, 7, L"Restart");
            AppendMenu(hMenu, MF_STRING, 2, L"Exit");
            
            SetForegroundWindow(hwnd);
            
            // TrackPopupMenu needs special handling for proper message delivery
            DWORD cmd = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, 
                                       pt.x, pt.y, 0, hwnd, nullptr);
            
            // Post a fake message to ensure menu closes properly
            PostMessage(hwnd, WM_NULL, 0, 0);
            
            if (cmd == 1) {
                mgr->ShowListWindow();
#ifdef HAVE_UIAUTOMATION
            } else if (cmd == 6) {
                if (!mgr->CopyFromFocusedControlViaUIA())
                    MessageBoxW(hwnd, L"No text could be read from the focused control.", L"clip2", MB_OK | MB_ICONINFORMATION);
#endif
            } else if (cmd == 7) {
                wchar_t exePath[MAX_PATH];
                GetModuleFileNameW(nullptr, exePath, MAX_PATH);
                ShellExecuteW(nullptr, L"open", exePath, nullptr, nullptr, SW_SHOWNORMAL);
                mgr->Stop();
                PostQuitMessage(0);
            } else if (cmd == 2) {
                mgr->Stop();
                PostQuitMessage(0);
            } else if (cmd == 3) {
                // Toggle startup with Windows
                mgr->SetStartupWithWindows(!startupEnabled);
            } else if (cmd == 4) {
                // Show settings dialog
                mgr->ShowSettingsDialog();
            } else if (cmd == 5) {
                mgr->ShowSnippetsManagerDialog();
            } else if (cmd >= 100 && cmd < 100 + (int)mgr->snippets.size()) {
                mgr->PasteSnippet(cmd - 100);
            }
            
            DestroyMenu(hMenu);
            return 0;
        }
        if (LOWORD(lParam) == WM_LBUTTONUP) {
            if (mgr->listVisible) {
                mgr->HideListWindow();
            } else {
                mgr->ShowListWindow();
            }
            return 0;
        }
        break;
        
    case WM_COMMAND:
        // Handle menu commands (fallback, though TPM_RETURNCMD should handle it)
        if (LOWORD(wParam) == 1) {
            mgr->ShowListWindow();
#ifdef HAVE_UIAUTOMATION
        } else if (LOWORD(wParam) == 6) {
            if (!mgr->CopyFromFocusedControlViaUIA())
                MessageBoxW(hwnd, L"No text could be read from the focused control.", L"clip2", MB_OK | MB_ICONINFORMATION);
#endif
        } else if (LOWORD(wParam) == 7) {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(nullptr, exePath, MAX_PATH);
            ShellExecuteW(nullptr, L"open", exePath, nullptr, nullptr, SW_SHOWNORMAL);
            mgr->Stop();
            PostQuitMessage(0);
        } else if (LOWORD(wParam) == 2) {
            mgr->Stop();
            PostQuitMessage(0);
        } else if (LOWORD(wParam) == 3) {
            // Toggle startup with Windows
            bool currentState = mgr->IsStartupWithWindows();
            mgr->SetStartupWithWindows(!currentState);
        }
        break;
        
    case WM_HOTKEY:
        if (wParam == ClipboardManager::HOTKEY_ID_OVERLAY) {
            DWORD now = GetTickCount();
            if (now - mgr->lastHotkeyTick < 250) return 0;  // Debounce (hook and RegisterHotKey may both fire)
            mgr->lastHotkeyTick = now;
            if (mgr->listVisible) {
                mgr->HideListWindow();
            } else {
                mgr->ShowListWindow();
            }
        }
#ifdef HAVE_UIAUTOMATION
        else if (wParam == ClipboardManager::HOTKEY_ID_COPY_FOCUSED) {
            if (!mgr->CopyFromFocusedControlViaUIA())
                MessageBoxW(hwnd, L"No text could be read from the focused control.", L"clip2", MB_OK | MB_ICONINFORMATION);
        }
#endif
        else if (wParam == ClipboardManager::HOTKEY_ID_PASTE_FOCUSED) {
            if (!mgr->PasteToFocusedControlWithoutClipboard(false))
                MessageBoxW(hwnd, L"Could not paste: no text in history, or the target could not receive keystrokes.", L"clip2", MB_OK | MB_ICONINFORMATION);
        } else if (wParam == ClipboardManager::HOTKEY_ID_PASTE_FOCUSED_CLIPBOARD) {
            if (!mgr->PasteToFocusedControlWithoutClipboard(true))
                MessageBoxW(hwnd, L"Could not paste: no text in history, or clipboard could not be set.", L"clip2", MB_OK | MB_ICONINFORMATION);
        }
        return 0;
    
    case WM_CLIPBOARD_HOTKEY:
        // Toggle list window visibility (from low-level keyboard hook)
        {
            DWORD now = GetTickCount();
            if (now - mgr->lastHotkeyTick < 250) return 0;  // Debounce
            mgr->lastHotkeyTick = now;
        }
        if (mgr->listVisible) {
            mgr->HideListWindow();
        } else {
            mgr->ShowListWindow();
        }
        return 0;
    
    case WM_SNIPPETS_OVERLAY_HOTKEY:
        {
            DWORD now = GetTickCount();
            if (now - mgr->lastHotkeyTick < 250) return 0;  // Debounce (same tick as overlay hotkey)
            mgr->lastHotkeyTick = now;
        }
        if (!mgr->listVisible) {
            mgr->ShowListWindow(true);
        } else if (mgr->snippetsMode) {
            mgr->HideListWindow();
        } else {
            mgr->SetListSnippetsMode(true);
        }
        return 0;
    
    case WM_DISMISS_OVERLAY:
        if (mgr->listVisible) {
            mgr->HideListWindow();
        }
        return 0;
    
    case WM_CLIPBOARDUPDATE:
        // Clipboard changed - instant notification from Windows
        if (!mgr->isPasting) {
            // Get current sequence number
            DWORD currentSequence = GetClipboardSequenceNumber();
            
            // If this matches the sequence number we set when pasting, ignore it
            // This prevents re-adding items we just pasted
            if (currentSequence != mgr->lastSequenceNumber) {
                // Play sound immediately when clipboard changes
                mgr->PlayClickSound();
                
                // Bypass copy blocks: capture clipboard content immediately before any app can clear/replace it
                mgr->TryCaptureClipboardImmediately();
                
                // Update sequence number
                mgr->lastSequenceNumber = currentSequence;
                
                // Post message to process clipboard asynchronously
                PostMessage(hwnd, WM_PROCESS_CLIPBOARD, 0, 0);
            }
        }
        return 0;
        
    case WM_PROCESS_CLIPBOARD:
        // Process clipboard change asynchronously
        mgr->ProcessClipboard();
        return 0;
    
    case WM_TIMER:
        // Check if foreground window is still our list window
        if (wParam == 1) { // Timer ID 1 for focus check
            if (mgr->listVisible && mgr->hwndList && IsWindowVisible(mgr->hwndList)) {
                HWND fgWindow = GetForegroundWindow();
                if (fgWindow != mgr->hwndList && fgWindow != mgr->hwndSearch) {
                    if (mgr->hwndPreview && fgWindow == mgr->hwndPreview) {
                        // Preview popup is still part of the overlay session
                    } else {
                        // Check if foreground window is a child of our list
                        bool isChild = false;
                        if (fgWindow != nullptr) {
                            isChild = IsChild(mgr->hwndList, fgWindow) != FALSE;
                        }
                        if (!isChild) {
                            mgr->HideListWindow();
                        }
                    }
                }
            } else {
                // Window is not visible or not supposed to be visible - stop timer
                KillTimer(hwnd, 1);
            }
        }
        return 0;
        
    case WM_CLOSE:
        mgr->HideListWindow();
        return 0;
        
    case WM_DESTROY:
        if (mgr) {
            mgr->isRunning = false;
        }
        PostQuitMessage(0);
        return 0;
    }
    
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK ClipboardManager::ListWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    
    // Handle preview window separately
    if (hwnd == mgr->hwndPreview) {
        switch (uMsg) {
        case WM_PAINT:
            {
                PAINTSTRUCT ps;
                HDC hdc = BeginPaint(hwnd, &ps);
                RECT rect;
                GetClientRect(hwnd, &rect);
                
                // Find which item is being previewed
                if (mgr->hoveredItemIndex >= 0 && mgr->hoveredItemIndex < (int)mgr->filteredIndices.size()) {
                    int actualIndex = mgr->filteredIndices[mgr->hoveredItemIndex];
                    const auto& item = mgr->clipboardHistory[actualIndex];
                    
                    // Draw background (themed)
                    { HBRUSH bgb = CreateSolidBrush(Theme5250::BG); FillRect(hdc, &rect, bgb); DeleteObject(bgb); }
                    
                    // Draw border - 5250 green
                    HPEN hPreviewBorderPen = CreatePen(PS_SOLID, 2, Theme5250::BORDER);
                    HPEN hOldPreviewPen = (HPEN)SelectObject(hdc, hPreviewBorderPen);
                    HBRUSH hOldPreviewBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
                    Rectangle(hdc, 1, 1, rect.right - 1, rect.bottom - 1);
                    SelectObject(hdc, hOldPreviewBrush);
                    SelectObject(hdc, hOldPreviewPen);
                    DeleteObject(hPreviewBorderPen);
                    
                    HBITMAP previewBmp = item->GetPreviewBitmap();
                    if (previewBmp) {
                        // Draw larger preview bitmap
                        HDC hdcMem = CreateCompatibleDC(hdc);
                        SelectObject(hdcMem, previewBmp);
                        
                        // Get bitmap size
                        BITMAP bm;
                        GetObject(previewBmp, sizeof(BITMAP), &bm);
                        
                        // Scale to fit preview window - use full available space with padding for border
                        int padding = 10; // Padding for rounded border
                        int maxW = rect.right - rect.left - (padding * 2);
                        int maxH = rect.bottom - rect.top - (padding * 2);
                        int srcW = bm.bmWidth;
                        int srcH = bm.bmHeight;
                        
                        // Calculate scale to fit within available space while maintaining aspect ratio
                        float scale = std::min((float)maxW / srcW, (float)maxH / srcH);
                        int dstW = (int)(srcW * scale);
                        int dstH = (int)(srcH * scale);
                        
                        // Center the image in the preview window
                        int offsetX = padding + (maxW - dstW) / 2;
                        int offsetY = padding + (maxH - dstH) / 2;
                        
                        // Use better scaling for quality
                        SetStretchBltMode(hdc, HALFTONE);
                        StretchBlt(hdc, offsetX, offsetY, dstW, dstH, hdcMem, 0, 0, srcW, srcH, SRCCOPY);
                        
                        DeleteDC(hdcMem);
                    } else if (item->isImage || item->isVideo) {
                        // Draw placeholder - Black background
                        RECT centerRect = rect;
                        centerRect.left += 50;
                        centerRect.top += 50;
                        centerRect.right -= 50;
                        centerRect.bottom -= 50;
                        HBRUSH brush = CreateSolidBrush(Theme5250::BG);
                        FillRect(hdc, &centerRect, brush);
                        DeleteObject(brush);
                        
                        SetTextColor(hdc, Theme5250::TXT);
                        SetBkMode(hdc, TRANSPARENT);
                        std::wstring text = item->isVideo ? L"Video Preview" : L"Image Preview";
                        DrawText(hdc, text.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                    }
                }
                
                EndPaint(hwnd, &ps);
                return 0;
            }
        case WM_MOUSELEAVE:
        case WM_KILLFOCUS:
            mgr->HidePreviewWindow();
            return 0;
        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
        }
    }
    
    switch (uMsg) {
    case WM_NCHITTEST: {
        // Allow window to be dragged by returning HTCAPTION for most areas
        // This makes the borderless window movable
        LRESULT hit = DefWindowProc(hwnd, uMsg, wParam, lParam);
        
        // If the default proc says it's on the client area, check if it's on an interactive element
        if (hit == HTCLIENT) {
            POINT pt = {LOWORD(lParam), HIWORD(lParam)};
            ScreenToClient(hwnd, &pt);
            
            RECT rect;
            GetClientRect(hwnd, &rect);
            
            // Check if click is on the title area (top 30 pixels) - allow dragging
            if (pt.y < 30) {
                return HTCAPTION;
            }
            
            // Check if click is on the search box area (y: 30-60) - don't allow dragging
            if (mgr->hwndSearch && IsWindowVisible(mgr->hwndSearch)) {
                if (pt.y >= 30 && pt.y <= 60) {
                    return HTCLIENT; // Let search box handle it
                }
            }
            
            // Check if click is on an item - let item handle it (for selection, paste, etc.)
            int itemHeight = 50;
            int yPos = 60;
            int startIndex = mgr->scrollOffset;
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            int endIndex = std::min(startIndex + mgr->itemsPerPage, listSize);
            
            for (int i = startIndex; i < endIndex; i++) {
                RECT itemRect = {5, yPos, rect.right - 5, yPos + itemHeight - 2};
                if (PtInRect(&itemRect, pt)) {
                    return HTCLIENT; // Let item handle it
                }
                yPos += itemHeight;
            }
            
            // If not on any interactive element, allow dragging
            return HTCAPTION;
        }
        
        return hit;
    }
    case WM_ERASEBKGND:
        if (hwnd == mgr->hwndList) return 1;  // fully painted in WM_PAINT (double-buffered): skip erase flicker
        break;

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdcWindow = BeginPaint(hwnd, &ps);

        RECT rect;
        GetClientRect(hwnd, &rect);

        // ---- Double buffering: render everything onto an off-screen bitmap first. ----
        HDC hdc = CreateCompatibleDC(hdcWindow);
        HBITMAP memBmp = CreateCompatibleBitmap(hdcWindow, rect.right, rect.bottom);
        HGDIOBJ oldMemBmp = SelectObject(hdc, memBmp);

        // Theme-derived accent tints.
        COLORREF accentSoft = BlendColor(Theme5250::BG, Theme5250::SEL_BG, 38);   // subtle row hover/header wash
        COLORREF dividerCol = BlendColor(Theme5250::BG, Theme5250::DIM, 150);     // faint separators

        // Layout constants. NOTE: content top (60), row height (50) and the search box
        // position (y=30) are part of the interaction contract used by hit-testing and
        // navigation; keep them in sync with GetItemAtPosition / WM_NCHITTEST.
        const int HEADER_H   = 28;
        const int SEARCH_TOP = 30, SEARCH_BOT = 56;
        const int CONTENT_TOP = 60;
        const int itemHeight = 50;
        const int FOOTER_H   = 24;
        const int listBottom = rect.bottom - FOOTER_H;

        SetBkMode(hdc, TRANSPARENT);

        // Background (themed).
        { HBRUSH bgb = CreateSolidBrush(Theme5250::BG); FillRect(hdc, &rect, bgb); DeleteObject(bgb); }

        HGDIOBJ hOldFont = SelectObject(hdc, GetOverlayFont());

        // ---- Header bar ----------------------------------------------------------
        {
            RECT headerRect = { 0, 0, rect.right, HEADER_H };
            HBRUSH hb = CreateSolidBrush(accentSoft);
            FillRect(hdc, &headerRect, hb);
            DeleteObject(hb);

            // App name (bold) on the left.
            SelectObject(hdc, GetOverlayFontBold());
            SetTextColor(hdc, Theme5250::TXT);
            RECT nameRect = { 12, 0, rect.right / 2, HEADER_H };
            DrawTextW(hdc, L"clip2", -1, &nameRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

            // Mode pill + live count on the right.
            int total = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size()
                                          : (int)mgr->filteredIndices.size();
            std::wstring modeText = mgr->snippetsMode ? L"SNIPPETS" : L"CLIPBOARD";
            std::wstring countText = std::to_wstring(total) + (total == 1 ? L" item" : L" items");
            if (!mgr->numberInput.empty()) countText = L"#" + mgr->numberInput;

            SelectObject(hdc, GetOverlayFontSmall());
            SIZE szMode{}, szCount{};
            GetTextExtentPoint32W(hdc, modeText.c_str(), (int)modeText.size(), &szMode);
            GetTextExtentPoint32W(hdc, countText.c_str(), (int)countText.size(), &szCount);

            int rightPad = 12;
            int countW = szCount.cx + 4;
            RECT countRect = { rect.right - rightPad - countW, 0, rect.right - rightPad, HEADER_H };
            SetTextColor(hdc, Theme5250::DIM == Theme5250::BG ? Theme5250::TXT : BlendColor(Theme5250::BG, Theme5250::TXT, 180));
            DrawTextW(hdc, countText.c_str(), -1, &countRect, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);

            // Pill behind the mode label.
            int pillPadX = 8;
            int pillRight = countRect.left - 10;
            int pillLeft = pillRight - (szMode.cx + pillPadX * 2);
            int pillTop = (HEADER_H - 16) / 2, pillBot = pillTop + 16;
            HPEN pen = CreatePen(PS_SOLID, 1, Theme5250::BORDER);
            HGDIOBJ op = SelectObject(hdc, pen);
            HGDIOBJ ob = SelectObject(hdc, GetStockObject(NULL_BRUSH));
            RoundRect(hdc, pillLeft, pillTop, pillRight, pillBot, 8, 8);
            SelectObject(hdc, ob);
            SelectObject(hdc, op);
            DeleteObject(pen);
            RECT pillRect = { pillLeft, pillTop, pillRight, pillBot };
            SetTextColor(hdc, Theme5250::TXT);
            DrawTextW(hdc, modeText.c_str(), -1, &pillRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            SelectObject(hdc, GetOverlayFont());
        }

        // Outer window border (thin, accent) for a crisp framed look.
        {
            HPEN hBorderPen = CreatePen(PS_SOLID, 1, Theme5250::BORDER);
            HGDIOBJ op = SelectObject(hdc, hBorderPen);
            HGDIOBJ ob = SelectObject(hdc, GetStockObject(NULL_BRUSH));
            Rectangle(hdc, 0, 0, rect.right, rect.bottom);
            SelectObject(hdc, ob);
            SelectObject(hdc, op);
            DeleteObject(hBorderPen);
        }

        // ---- Search field --------------------------------------------------------
        if (mgr->hwndSearch && IsWindowVisible(mgr->hwndSearch)) {
            bool searchFocused = (GetFocus() == mgr->hwndSearch);
            RECT searchRect = { 6, SEARCH_TOP, rect.right - 6, SEARCH_BOT };
            HPEN hPen = CreatePen(PS_SOLID, 1, searchFocused ? Theme5250::TXT : dividerCol);
            HGDIOBJ op = SelectObject(hdc, hPen);
            HGDIOBJ ob = SelectObject(hdc, GetStockObject(NULL_BRUSH));
            RoundRect(hdc, searchRect.left, searchRect.top, searchRect.right, searchRect.bottom, 6, 6);
            SelectObject(hdc, ob);
            SelectObject(hdc, op);
            DeleteObject(hPen);
        }

        // ---- Scroll / paging math ------------------------------------------------
        int totalItems = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
        int visibleHeight = listBottom - CONTENT_TOP;
        mgr->itemsPerPage = std::max(1, visibleHeight / itemHeight);
        int maxScroll = std::max(0, totalItems - mgr->itemsPerPage);
        if (mgr->scrollOffset > maxScroll) mgr->scrollOffset = maxScroll;
        if (mgr->scrollOffset < 0) mgr->scrollOffset = 0;

        int yPos = CONTENT_TOP;
        int startIndex = mgr->scrollOffset;
        int endIndex = std::min(startIndex + mgr->itemsPerPage, totalItems);
        int rowRight = rect.right - 9;  // leave room for scrollbar

        // Empty-state hint.
        if (totalItems == 0) {
            SetTextColor(hdc, BlendColor(Theme5250::BG, Theme5250::DIM, 220));
            RECT emptyRect = { 0, CONTENT_TOP, rect.right, listBottom };
            std::wstring msg = mgr->searchText.empty()
                ? (mgr->snippetsMode ? L"No snippets yet" : L"Clipboard history is empty")
                : L"No matches";
            DrawTextW(hdc, msg.c_str(), -1, &emptyRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }

        if (mgr->snippetsMode) {
            for (int i = startIndex; i < endIndex && yPos + itemHeight <= listBottom; i++) {
                if (i >= (int)mgr->filteredSnippetIndices.size()) break;
                int actualIndex = mgr->filteredSnippetIndices[i];
                const auto& snip = mgr->snippets[actualIndex];

                bool isSelected = (i == mgr->selectedIndex && mgr->selectedIndex >= 0);
                bool isHover = (i == mgr->hoveredItemIndex);
                RECT rowRect = { 6, yPos + 2, rowRight, yPos + itemHeight - 2 };

                if (isSelected) {
                    HBRUSH b = CreateSolidBrush(Theme5250::SEL_BG);
                    HRGN rgn = CreateRoundRectRgn(rowRect.left, rowRect.top, rowRect.right, rowRect.bottom, 8, 8);
                    FillRgn(hdc, rgn, b);
                    DeleteObject(rgn);
                    DeleteObject(b);
                } else if (isHover) {
                    HBRUSH b = CreateSolidBrush(accentSoft);
                    HRGN rgn = CreateRoundRectRgn(rowRect.left, rowRect.top, rowRect.right, rowRect.bottom, 8, 8);
                    FillRgn(hdc, rgn, b);
                    DeleteObject(rgn);
                    DeleteObject(b);
                }

                COLORREF fg = isSelected ? Theme5250::SEL_FG : Theme5250::TXT;
                COLORREF meta = isSelected ? Theme5250::SEL_FG : BlendColor(Theme5250::BG, Theme5250::TXT, 150);

                RECT numRect = { rowRect.left + 8, rowRect.top, rowRect.left + 40, rowRect.bottom };
                SetTextColor(hdc, meta);
                DrawTextW(hdc, (std::to_wstring(i + 1)).c_str(), -1, &numRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

                RECT nameRect = { rowRect.left + 46, rowRect.top + 6, rowRect.right - 12, rowRect.top + 24 };
                SetTextColor(hdc, fg);
                SelectObject(hdc, GetOverlayFontBold());
                DrawTextW(hdc, snip.name.c_str(), -1, &nameRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);
                SelectObject(hdc, GetOverlayFont());

                std::wstring dispContent = snip.contentPlain.empty() ? snip.content : snip.contentPlain;
                if (!dispContent.empty()) {
                    for (auto& ch : dispContent) if (ch == L'\r' || ch == L'\n' || ch == L'\t') ch = L' ';
                    RECT contentRect = { rowRect.left + 46, rowRect.top + 24, rowRect.right - 12, rowRect.bottom - 4 };
                    SetTextColor(hdc, meta);
                    SelectObject(hdc, GetOverlayFontSmall());
                    DrawTextW(hdc, dispContent.c_str(), -1, &contentRect, DT_LEFT | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);
                    SelectObject(hdc, GetOverlayFont());
                }
                yPos += itemHeight;
            }
        } else {
            for (int i = startIndex; i < endIndex && yPos + itemHeight <= listBottom; i++) {
                if (i >= (int)mgr->filteredIndices.size()) break;
                int actualIndex = mgr->filteredIndices[i];
                const auto& item = mgr->clipboardHistory[actualIndex];

                bool isMultiSelected = mgr->multiSelectedIndices.find(i) != mgr->multiSelectedIndices.end();
                bool isSelected = (i == mgr->selectedIndex && mgr->selectedIndex >= 0);
                bool isHover = (i == mgr->hoveredItemIndex);
                bool hot = isMultiSelected || isSelected;
                RECT rowRect = { 6, yPos + 2, rowRight, yPos + itemHeight - 2 };

                if (hot) {
                    HBRUSH b = CreateSolidBrush(Theme5250::SEL_BG);
                    HRGN rgn = CreateRoundRectRgn(rowRect.left, rowRect.top, rowRect.right, rowRect.bottom, 8, 8);
                    FillRgn(hdc, rgn, b);
                    DeleteObject(rgn);
                    DeleteObject(b);
                } else if (isHover) {
                    HBRUSH b = CreateSolidBrush(accentSoft);
                    HRGN rgn = CreateRoundRectRgn(rowRect.left, rowRect.top, rowRect.right, rowRect.bottom, 8, 8);
                    FillRgn(hdc, rgn, b);
                    DeleteObject(rgn);
                    DeleteObject(b);
                }

                COLORREF fg = hot ? Theme5250::SEL_FG : Theme5250::TXT;
                COLORREF meta = hot ? Theme5250::SEL_FG : BlendColor(Theme5250::BG, Theme5250::TXT, 150);

                // Pinned accent bar at the very left of the row.
                if (item->pinned) {
                    RECT pinBar = { rowRect.left, rowRect.top + 4, rowRect.left + 3, rowRect.bottom - 4 };
                    HBRUSH pinBrush = CreateSolidBrush(fg);
                    FillRect(hdc, &pinBar, pinBrush);
                    DeleteObject(pinBrush);
                }

                int contentLeft = rowRect.left + 10;

                // Thumbnail or a type badge.
                if (item->thumbnail) {
                    int thumbSize = 36;
                    int thumbY = yPos + (itemHeight - thumbSize) / 2;
                    HDC hdcMem = CreateCompatibleDC(hdc);
                    HGDIOBJ ob = SelectObject(hdcMem, item->thumbnail);
                    SetStretchBltMode(hdc, HALFTONE);
                    BitBlt(hdc, contentLeft, thumbY, thumbSize, thumbSize, hdcMem, 0, 0, SRCCOPY);
                    SelectObject(hdcMem, ob);
                    DeleteDC(hdcMem);
                    contentLeft += thumbSize + 10;
                } else {
                    // Letter badge in a rounded box indicating the item type.
                    const wchar_t* badge = L"T";
                    if (item->isImage) badge = L"IMG";
                    else if (item->isVideo) badge = L"VID";
                    else if (item->fileType == L"Files") badge = L"FILE";
                    else if (item->fileType == L"Text") badge = L"T";
                    else badge = L"?";
                    int bW = 34, bH = 22;
                    int bx = contentLeft, by = yPos + (itemHeight - bH) / 2;
                    HPEN pen = CreatePen(PS_SOLID, 1, hot ? Theme5250::SEL_FG : Theme5250::BORDER);
                    HGDIOBJ op = SelectObject(hdc, pen);
                    HGDIOBJ ob = SelectObject(hdc, GetStockObject(NULL_BRUSH));
                    RoundRect(hdc, bx, by, bx + bW, by + bH, 6, 6);
                    SelectObject(hdc, ob);
                    SelectObject(hdc, op);
                    DeleteObject(pen);
                    RECT badgeRect = { bx, by, bx + bW, by + bH };
                    SetTextColor(hdc, meta);
                    SelectObject(hdc, GetOverlayFontSmall());
                    DrawTextW(hdc, badge, -1, &badgeRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                    SelectObject(hdc, GetOverlayFont());
                    contentLeft += bW + 10;
                }

                // Index number.
                RECT numRect = { contentLeft, rowRect.top, contentLeft + 28, rowRect.bottom };
                SetTextColor(hdc, meta);
                DrawTextW(hdc, (std::to_wstring(i + 1)).c_str(), -1, &numRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
                contentLeft += 34;

                // Metadata (format + relative time), right-aligned.
                std::wstring info = item->formatName + L"  \x2022  " + RelativeTimeString(item->timestamp);
                SelectObject(hdc, GetOverlayFontSmall());
                SIZE szInfo{};
                GetTextExtentPoint32W(hdc, info.c_str(), (int)info.size(), &szInfo);
                int infoTextW = std::min((int)szInfo.cx, 210);
                int infoBoxW = infoTextW + 14;  // include right padding so text isn't clipped
                RECT infoRect = { rowRect.right - infoBoxW, rowRect.top, rowRect.right - 8, rowRect.bottom };
                SetTextColor(hdc, meta);
                DrawTextW(hdc, info.c_str(), -1, &infoRect, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                SelectObject(hdc, GetOverlayFont());

                // Preview text fills the space between number and metadata.
                std::wstring previewText = item->preview;
                for (auto& ch : previewText) if (ch == L'\r' || ch == L'\n' || ch == L'\t') ch = L' ';
                RECT previewRect = { contentLeft, rowRect.top, rowRect.right - infoBoxW - 12, rowRect.bottom };
                SetTextColor(hdc, fg);
                DrawTextW(hdc, previewText.c_str(), -1, &previewRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

                yPos += itemHeight;
            }
        }

        // ---- Scrollbar indicator -------------------------------------------------
        if (totalItems > mgr->itemsPerPage) {
            int trackTop = CONTENT_TOP + 2, trackBot = listBottom - 2;
            int trackH = trackBot - trackTop;
            int trackX = rect.right - 6;
            if (trackH > 10) {
                HBRUSH tb = CreateSolidBrush(dividerCol);
                RECT track = { trackX, trackTop, trackX + 3, trackBot };
                FillRect(hdc, &track, tb);
                DeleteObject(tb);
                int thumbH = std::max(20, trackH * mgr->itemsPerPage / totalItems);
                int thumbY = trackTop + (maxScroll > 0 ? (trackH - thumbH) * mgr->scrollOffset / maxScroll : 0);
                HBRUSH thb = CreateSolidBrush(Theme5250::BORDER);
                RECT thumb = { trackX, thumbY, trackX + 3, thumbY + thumbH };
                FillRect(hdc, &thumb, thb);
                DeleteObject(thb);
            }
        }

        // ---- Footer status / hint bar --------------------------------------------
        {
            RECT footRect = { 0, rect.bottom - FOOTER_H, rect.right, rect.bottom };
            HBRUSH fb = CreateSolidBrush(accentSoft);
            FillRect(hdc, &footRect, fb);
            DeleteObject(fb);
            // Divider line above footer.
            HPEN dp = CreatePen(PS_SOLID, 1, dividerCol);
            HGDIOBJ op = SelectObject(hdc, dp);
            MoveToEx(hdc, 1, footRect.top, nullptr);
            LineTo(hdc, rect.right - 1, footRect.top);
            SelectObject(hdc, op);
            DeleteObject(dp);

            SelectObject(hdc, GetOverlayFontSmall());
            SetTextColor(hdc, BlendColor(Theme5250::BG, Theme5250::TXT, 170));
            RECT hintRect = { 12, footRect.top, rect.right - 12, footRect.bottom };
            const wchar_t* hint = mgr->snippetsMode
                ? L"Enter paste  \x2022  Ctrl+Left clipboard  \x2022  Ctrl+F search  \x2022  Esc close"
                : L"1-9 quick  \x2022  Enter paste  \x2022  Ctrl+Right snippets  \x2022  Ctrl+F search  \x2022  Esc close";
            DrawTextW(hdc, hint, -1, &hintRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            SelectObject(hdc, GetOverlayFont());
        }

        // Blit the finished frame to the window in one shot (flicker-free).
        BitBlt(hdcWindow, 0, 0, rect.right, rect.bottom, hdc, 0, 0, SRCCOPY);

        SelectObject(hdc, hOldFont);
        SelectObject(hdc, oldMemBmp);
        DeleteObject(memBmp);
        DeleteDC(hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    
    case WM_KEYDOWN: {
        // Ctrl+Right = Snippets mode, Ctrl+Left = Clipboard mode
        bool ctrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrlPressed && (wParam == VK_RIGHT || wParam == VK_LEFT)) {
            mgr->SetListSnippetsMode(wParam == VK_RIGHT);
            return 0;
        }
        // Handle Enter key first
        if (wParam == VK_RETURN) {
            if (mgr->snippetsMode) {
                if (!mgr->numberInput.empty()) {
                    try {
                        int itemNumber = std::stoi(mgr->numberInput);
                        int index = itemNumber - 1;
                        if (index >= 0 && index < (int)mgr->filteredSnippetIndices.size()) {
                            mgr->PasteSnippet(mgr->filteredSnippetIndices[index]);
                            mgr->numberInput.clear();
                            mgr->HideListWindow();
                        } else {
                            mgr->numberInput.clear();
                            mgr->UpdateListWindow();
                        }
                    } catch (...) {
                        mgr->numberInput.clear();
                        mgr->UpdateListWindow();
                    }
                } else if (mgr->selectedIndex >= 0 && mgr->selectedIndex < (int)mgr->filteredSnippetIndices.size()) {
                    mgr->PasteSnippet(mgr->filteredSnippetIndices[mgr->selectedIndex]);
                    mgr->HideListWindow();
                }
            } else {
                // Clipboard mode
                if (!mgr->numberInput.empty()) {
                    try {
                        int itemNumber = std::stoi(mgr->numberInput);
                        int index = itemNumber - 1;
                        if (index >= 0 && index < (int)mgr->filteredIndices.size()) {
                            mgr->PasteItem(index);
                            mgr->numberInput.clear();
                            mgr->HideListWindow();
                        } else {
                            mgr->numberInput.clear();
                            mgr->UpdateListWindow();
                        }
                    } catch (...) {
                        mgr->numberInput.clear();
                        mgr->UpdateListWindow();
                    }
                } else {
                    if (!mgr->multiSelectedIndices.empty()) {
                        mgr->PasteMultipleItems();
                        mgr->HideListWindow();
                    } else if (mgr->selectedIndex >= 0 && mgr->selectedIndex < (int)mgr->filteredIndices.size()) {
                        mgr->PasteItem(mgr->selectedIndex);
                        mgr->HideListWindow();
                    }
                }
            }
            return 0;
        }
        // Handle Backspace (clipboard mode only - snippets mode forwards to search)
        if (wParam == VK_BACK && !mgr->snippetsMode) {
            if (!mgr->numberInput.empty()) {
                mgr->numberInput.pop_back();
                mgr->UpdateListWindow();
            }
            return 0;
        }
        // Quick paste for single digits 1-9 (clipboard mode only; snippets mode uses search + Enter)
        if (!mgr->snippetsMode && wParam >= 0x31 && wParam <= 0x39 && mgr->numberInput.empty()) {
            int index = (wParam - 0x31);
            if (index < (int)mgr->filteredIndices.size()) {
                mgr->PasteItem(index);
                mgr->HideListWindow();
            }
            return 0;
        }
        // In snippets mode: forward Backspace to search box (printable keys go via WM_CHAR)
        if (mgr->snippetsMode && mgr->hwndSearch) {
            HWND focusedWindow = GetFocus();
            if (focusedWindow != mgr->hwndSearch && wParam == VK_BACK) {
                SetFocus(mgr->hwndSearch);
                PostMessage(mgr->hwndSearch, WM_KEYDOWN, VK_BACK, lParam);
                return 0;
            }
        }
        // Handle Ctrl+F to focus search
        if (wParam == 'F' && (GetAsyncKeyState(VK_CONTROL) & 0x8000)) {
            if (mgr->hwndSearch) {
                SetFocus(mgr->hwndSearch);
                // Select all text in search box
                SendMessage(mgr->hwndSearch, EM_SETSEL, 0, -1);
            }
            return 0;
        }
        // For number keys, don't return 0 - let WM_CHAR handle them
        if (wParam >= 0x30 && wParam <= 0x39) {
            // Let it fall through to WM_CHAR
            break;
        }
        if (wParam == VK_ESCAPE) {
            mgr->numberInput.clear();
            mgr->HideListWindow();
            return 0;
        }
        // Handle Delete key - only in clipboard mode
        if (wParam == VK_DELETE && !mgr->snippetsMode) {
            HWND focusedWindow = GetFocus();
            if (focusedWindow != mgr->hwndSearch && mgr->selectedIndex >= 0 && mgr->selectedIndex < (int)mgr->filteredIndices.size()) {
                mgr->DeleteItem(mgr->selectedIndex);
                return 0;
            }
        }
        // Navigate with arrow keys (only if search box doesn't have focus)
        {
            HWND focusedWindow = GetFocus();
            if (focusedWindow != mgr->hwndSearch) {
                int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
                bool isShiftPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                if (mgr->snippetsMode) isShiftPressed = false;  // No multi-select for snippets
                
                if (wParam == VK_UP) {
                    // Move selection up
                    if (mgr->selectedIndex > 0) {
                        int newIndex = mgr->selectedIndex - 1;
                        
                        if (isShiftPressed) {
                            // Shift+Up: Extend multi-selection
                            if (mgr->multiSelectAnchor < 0) {
                                // Start new range selection from current position
                                mgr->multiSelectAnchor = mgr->selectedIndex;
                                mgr->multiSelectedIndices.clear();
                                mgr->multiSelectedIndices.insert(mgr->selectedIndex);
                            }
                            // Extend selection to new index
                            int startIdx = std::min(mgr->multiSelectAnchor, newIndex);
                            int endIdx = std::max(mgr->multiSelectAnchor, newIndex);
                            mgr->multiSelectedIndices.clear();
                            for (int idx = startIdx; idx <= endIdx; idx++) {
                                mgr->multiSelectedIndices.insert(idx);
                            }
                            mgr->selectedIndex = newIndex;
                        } else {
                            // Regular Up: Clear multi-selection and move single selection
                            mgr->ClearMultiSelection();
                            mgr->multiSelectAnchor = -1;
                            mgr->selectedIndex = newIndex;
                        }
                        
                        // Ensure selected item is visible
                        if (mgr->selectedIndex < mgr->scrollOffset) {
                            mgr->scrollOffset = mgr->selectedIndex;
                        }
                        mgr->UpdateListWindow();
                    }
                    return 0;
                }
                if (wParam == VK_DOWN) {
                    int maxIndex = listSize - 1;
                    if (mgr->selectedIndex < maxIndex) {
                        int newIndex = mgr->selectedIndex + 1;
                        
                        if (isShiftPressed) {
                            // Shift+Down: Extend multi-selection
                            if (mgr->multiSelectAnchor < 0) {
                                // Start new range selection from current position
                                mgr->multiSelectAnchor = mgr->selectedIndex;
                                mgr->multiSelectedIndices.clear();
                                mgr->multiSelectedIndices.insert(mgr->selectedIndex);
                            }
                            // Extend selection to new index
                            int startIdx = std::min(mgr->multiSelectAnchor, newIndex);
                            int endIdx = std::max(mgr->multiSelectAnchor, newIndex);
                            mgr->multiSelectedIndices.clear();
                            for (int idx = startIdx; idx <= endIdx; idx++) {
                                mgr->multiSelectedIndices.insert(idx);
                            }
                            mgr->selectedIndex = newIndex;
                        } else {
                            // Regular Down: Clear multi-selection and move single selection
                            mgr->ClearMultiSelection();
                            mgr->multiSelectAnchor = -1;
                            mgr->selectedIndex = newIndex;
                        }
                        
                        int maxScroll = std::max(0, listSize - mgr->itemsPerPage);
                        if (mgr->selectedIndex >= mgr->scrollOffset + mgr->itemsPerPage) {
                            mgr->scrollOffset = mgr->selectedIndex - mgr->itemsPerPage + 1;
                            if (mgr->scrollOffset > maxScroll) {
                                mgr->scrollOffset = maxScroll;
                            }
                        }
                        mgr->UpdateListWindow();
                    }
                    return 0;
                }
            }
        }
        if (wParam == VK_PRIOR) {
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            bool isShiftPressed = mgr->snippetsMode ? false : ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0);
            int newIndex = std::max(0, mgr->scrollOffset - mgr->itemsPerPage);
            mgr->scrollOffset = newIndex;
            
            if (isShiftPressed && mgr->multiSelectAnchor >= 0) {
                // Extend multi-selection
                int startIdx = std::min(mgr->multiSelectAnchor, newIndex);
                int endIdx = std::max(mgr->multiSelectAnchor, newIndex);
                mgr->multiSelectedIndices.clear();
                for (int idx = startIdx; idx <= endIdx; idx++) {
                    mgr->multiSelectedIndices.insert(idx);
                }
            } else if (!isShiftPressed) {
                mgr->ClearMultiSelection();
            }
            mgr->selectedIndex = newIndex;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_NEXT) {
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            bool isShiftPressed = mgr->snippetsMode ? false : ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0);
            int maxScroll = std::max(0, listSize - mgr->itemsPerPage);
            int newIndex = std::min(maxScroll, mgr->scrollOffset + mgr->itemsPerPage);
            mgr->scrollOffset = newIndex;
            
            if (isShiftPressed && mgr->multiSelectAnchor >= 0) {
                // Extend multi-selection
                int startIdx = std::min(mgr->multiSelectAnchor, newIndex);
                int endIdx = std::max(mgr->multiSelectAnchor, newIndex);
                mgr->multiSelectedIndices.clear();
                for (int idx = startIdx; idx <= endIdx; idx++) {
                    mgr->multiSelectedIndices.insert(idx);
                }
            } else if (!isShiftPressed) {
                mgr->ClearMultiSelection();
            }
            mgr->selectedIndex = newIndex;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_HOME) {
            bool isShiftPressed = mgr->snippetsMode ? false : ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0);
            mgr->scrollOffset = 0;
            int newIndex = 0;
            
            if (isShiftPressed && mgr->multiSelectAnchor >= 0) {
                // Extend multi-selection to start
                int startIdx = 0;
                int endIdx = std::max(mgr->multiSelectAnchor, mgr->selectedIndex);
                mgr->multiSelectedIndices.clear();
                for (int idx = startIdx; idx <= endIdx; idx++) {
                    mgr->multiSelectedIndices.insert(idx);
                }
            } else if (!isShiftPressed) {
                mgr->ClearMultiSelection();
            }
            mgr->selectedIndex = newIndex;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_END) {
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            bool isShiftPressed = mgr->snippetsMode ? false : ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0);
            int maxScroll = std::max(0, listSize - mgr->itemsPerPage);
            mgr->scrollOffset = maxScroll;
            int maxIndex = listSize - 1;
            int newIndex = maxIndex >= 0 ? maxIndex : 0;
            
            if (isShiftPressed && mgr->multiSelectAnchor >= 0) {
                // Extend multi-selection to end
                int startIdx = std::min(mgr->multiSelectAnchor, mgr->selectedIndex);
                int endIdx = newIndex;
                mgr->multiSelectedIndices.clear();
                for (int idx = startIdx; idx <= endIdx; idx++) {
                    mgr->multiSelectedIndices.insert(idx);
                }
            } else if (!isShiftPressed) {
                mgr->ClearMultiSelection();
            }
            mgr->selectedIndex = newIndex;
            mgr->UpdateListWindow();
            return 0;
        }
        break;
    }
        
    case WM_CHAR:
        // In snippets mode: all typing goes to search box - show what user writes, Enter pastes match
        if (mgr->snippetsMode && mgr->hwndSearch) {
            if (mgr->ignoreNextSChar && (wParam == 's' || wParam == 'S')) {
                mgr->ignoreNextSChar = false;
                return 0;  // Consume the second 's' from "ss" - don't add to search
            }
            mgr->ignoreNextSChar = false;
            HWND focusedWindow = GetFocus();
            if (focusedWindow != mgr->hwndSearch) {
                SetFocus(mgr->hwndSearch);
            }
            PostMessage(mgr->hwndSearch, WM_CHAR, wParam, lParam);
            return 0;
        }
        // Clipboard mode: numeric input for quick paste
        if (wParam >= '0' && wParam <= '9') {
            if (mgr->numberInput.length() < 4) {
                mgr->numberInput += (wchar_t)wParam;
                mgr->UpdateListWindow();
            }
            return 0;
        }
        break;
        
    case WM_VSCROLL:
        {
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            int maxScroll = std::max(0, listSize - mgr->itemsPerPage);
            int oldOffset = mgr->scrollOffset;
            
            switch (LOWORD(wParam)) {
            case SB_LINEUP:
                mgr->scrollOffset = std::max(0, mgr->scrollOffset - 1);
                break;
            case SB_LINEDOWN:
                mgr->scrollOffset = std::min(maxScroll, mgr->scrollOffset + 1);
                break;
            case SB_PAGEUP:
                mgr->scrollOffset = std::max(0, mgr->scrollOffset - mgr->itemsPerPage);
                break;
            case SB_PAGEDOWN:
                mgr->scrollOffset = std::min(maxScroll, mgr->scrollOffset + mgr->itemsPerPage);
                break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION:
                {
                    // Get scroll position from message (since scrollbar is hidden)
                    mgr->scrollOffset = HIWORD(wParam);
                    if (mgr->scrollOffset > maxScroll) mgr->scrollOffset = maxScroll;
                    if (mgr->scrollOffset < 0) mgr->scrollOffset = 0;
                }
                break;
            case SB_TOP:
                mgr->scrollOffset = 0;
                break;
            case SB_BOTTOM:
                mgr->scrollOffset = maxScroll;
                break;
            }
            
            if (mgr->scrollOffset != oldOffset) {
                mgr->UpdateListWindow();
            }
            return 0;
        }
        
    case WM_MOUSEMOVE:
        {
            // Track mouse movement for hover preview
            POINT pt;
            pt.x = LOWORD(lParam);
            pt.y = HIWORD(lParam);
            
            int itemIndex = mgr->GetItemAtPosition(pt.x, pt.y);
            bool hoverChanged = (mgr->hoveredItemIndex != itemIndex);

            // Update the row hover highlight and repaint only when it changes.
            if (hoverChanged) {
                mgr->hoveredItemIndex = itemIndex;
                mgr->UpdateListWindow();
            }

            if (mgr->snippetsMode) {
                mgr->HidePreviewWindow();
            } else if (itemIndex >= 0 && itemIndex < (int)mgr->filteredIndices.size()) {
                int actualIndex = mgr->filteredIndices[itemIndex];
                const auto& item = mgr->clipboardHistory[actualIndex];
                if (item->isImage || item->isVideo) {
                    if (hoverChanged) {
                        POINT screenPt = pt;
                        ClientToScreen(hwnd, &screenPt);
                        mgr->ShowPreviewWindow(itemIndex, screenPt.x, screenPt.y);
                    }
                } else {
                    mgr->HidePreviewWindow();
                }
            } else {
                mgr->HidePreviewWindow();
            }
            
            // Track mouse leave
            TRACKMOUSEEVENT tme = {};
            tme.cbSize = sizeof(TRACKMOUSEEVENT);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hwnd;
            TrackMouseEvent(&tme);
            
            break;
        }
        
    case WM_MOUSELEAVE:
        mgr->HidePreviewWindow();
        if (mgr->hoveredItemIndex != -1) {
            mgr->hoveredItemIndex = -1;
            mgr->UpdateListWindow();
        }
        break;
        
    case WM_LBUTTONDOWN:
        {
            POINT pt;
            pt.x = LOWORD(lParam);
            pt.y = HIWORD(lParam);
            int clickedIndex = mgr->GetItemAtPosition(pt.x, pt.y);
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            
            if (clickedIndex >= 0 && clickedIndex < listSize) {
                if (mgr->snippetsMode) {
                    mgr->PasteSnippet(mgr->filteredSnippetIndices[clickedIndex]);
                    mgr->HideListWindow();
                } else {
                    bool isCtrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                    if (isCtrlPressed) {
                        mgr->ToggleMultiSelect(clickedIndex);
                        mgr->selectedIndex = clickedIndex;
                        mgr->multiSelectAnchor = -1;
                    } else {
                        mgr->ClearMultiSelection();
                        mgr->selectedIndex = clickedIndex;
                        mgr->UpdateListWindow();
                    }
                }
            }
            break;
        }
        
    case WM_LBUTTONDBLCLK:
        {
            POINT pt;
            pt.x = LOWORD(lParam);
            pt.y = HIWORD(lParam);
            int clickedIndex = mgr->GetItemAtPosition(pt.x, pt.y);
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            
            if (clickedIndex >= 0 && clickedIndex < listSize) {
                if (mgr->snippetsMode) {
                    mgr->PasteSnippet(mgr->filteredSnippetIndices[clickedIndex]);
                } else if (!mgr->multiSelectedIndices.empty()) {
                    mgr->PasteMultipleItems();
                } else {
                    mgr->PasteItem(clickedIndex);
                }
                mgr->numberInput.clear();
                mgr->HidePreviewWindow();
                mgr->HideListWindow();
            }
            return 0;
        }
        
    case WM_COMMAND:
        // Handle search edit control text change
        if (HIWORD(wParam) == EN_CHANGE && (HWND)lParam == mgr->hwndSearch) {
            int length = GetWindowTextLength(mgr->hwndSearch);
            if (length > 0) {
                wchar_t* buffer = new wchar_t[length + 1];
                GetWindowText(mgr->hwndSearch, buffer, length + 1);
                mgr->searchText = buffer;
                delete[] buffer;
            } else {
                mgr->searchText.clear();
            }
            if (mgr->snippetsMode) {
                mgr->FilterSnippets();
            } else {
                mgr->FilterItems();
            }
            return 0;
        }
        break;
        
    case WM_RBUTTONDOWN:
    case WM_CONTEXTMENU:
        {
            // Show context menu on right-click
            POINT pt;
            POINT clientPt;
            if (uMsg == WM_RBUTTONDOWN) {
                pt.x = LOWORD(lParam);
                pt.y = HIWORD(lParam);
                clientPt = pt;
                ClientToScreen(hwnd, &pt);
            } else {
                pt.x = LOWORD(lParam);
                pt.y = HIWORD(lParam);
                if (pt.x == -1 && pt.y == -1) {
                    // Menu was triggered by keyboard (Shift+F10)
                    RECT rect;
                    GetClientRect(hwnd, &rect);
                    pt.x = rect.right / 2;
                    pt.y = rect.bottom / 2;
                    clientPt.x = pt.x;
                    clientPt.y = pt.y;
                    ClientToScreen(hwnd, &pt);
                } else {
                    clientPt = pt;
                    ScreenToClient(hwnd, &clientPt);
                }
            }
            
            int clickedItemIndex = mgr->GetItemAtPosition(clientPt.x, clientPt.y);
            if (clickedItemIndex < 0 && uMsg == WM_CONTEXTMENU) {
                clickedItemIndex = mgr->selectedIndex;
            }
            
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            HMENU hMenu = CreatePopupMenu();
            
            if (mgr->snippetsMode) {
                if (clickedItemIndex >= 0 && clickedItemIndex < listSize) {
                    AppendMenu(hMenu, MF_STRING, 102, L"Paste");
                    AppendMenu(hMenu, MF_STRING, 105, L"Delete Snippet");
                    AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
                }
                AppendMenu(hMenu, MF_STRING, 104, L"Add Snippet...");
                AppendMenu(hMenu, MF_STRING, 103, L"Manage Snippets...");
            } else {
            bool isText = false;
            if (clickedItemIndex >= 0 && clickedItemIndex < listSize) {
                int actualIndex = mgr->filteredIndices[clickedItemIndex];
                isText = mgr->IsTextItem(actualIndex);
            }
            
            if (isText && clickedItemIndex >= 0 && clickedItemIndex < listSize) {
                HMENU hTransformMenu = CreatePopupMenu();
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_UPPERCASE, L"Uppercase");
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_LOWERCASE, L"Lowercase");
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_TITLE_CASE, L"Title Case");
                AppendMenu(hTransformMenu, MF_SEPARATOR, 0, nullptr);
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_REMOVE_LINE_BREAKS, L"Remove Line Breaks");
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_TRIM_WHITESPACE, L"Trim Whitespace");
                AppendMenu(hTransformMenu, MF_SEPARATOR, 0, nullptr);
                AppendMenu(hTransformMenu, MF_STRING, TRANSFORM_PLAIN_TEXT, L"Plain Text (Remove Formatting)");
                
                AppendMenu(hMenu, MF_STRING | MF_POPUP, (UINT_PTR)hTransformMenu, L"Transform Text");
                AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            }
            
            if (clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                AppendMenu(hMenu, MF_STRING, 106, L"Keystroke paste");
            }
            
            // Quick actions (clipboard items only).
            if (clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                int actualIndex = mgr->filteredIndices[clickedItemIndex];
                bool isPinned = (actualIndex >= 0 && actualIndex < (int)mgr->clipboardHistory.size() &&
                                 mgr->clipboardHistory[actualIndex] && mgr->clipboardHistory[actualIndex]->pinned);
                bool isImageItem = (actualIndex >= 0 && actualIndex < (int)mgr->clipboardHistory.size() &&
                                    mgr->clipboardHistory[actualIndex] && mgr->clipboardHistory[actualIndex]->isImage);
                AppendMenu(hMenu, MF_STRING, 110, isPinned ? L"Unpin" : L"Pin");
                if (isText) {
                    AppendMenu(hMenu, MF_STRING, 111, L"Copy as plain text");
                    // Show "Open URL" only when the item text looks like an http/https URL.
                    std::wstring t = GetPlainTextForDirectPaste(mgr->clipboardHistory[actualIndex].get());
                    size_t a = t.find_first_not_of(L" \t\r\n");
                    if (a != std::wstring::npos) {
                        std::wstring u = t.substr(a);
                        if (u.rfind(L"http://", 0) == 0 || u.rfind(L"https://", 0) == 0)
                            AppendMenu(hMenu, MF_STRING, 112, L"Open URL");
                    }
                }
                if (isImageItem) {
                    AppendMenu(hMenu, MF_STRING, 113, L"Save image to file...");
                }
                AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            }
            
            if (clickedItemIndex >= 0 && clickedItemIndex < listSize) {
                AppendMenu(hMenu, MF_STRING, 101, L"Delete");
                AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            }
            
            AppendMenu(hMenu, MF_STRING, 100, L"Clear List");
            }
            
            SetForegroundWindow(hwnd);
            
            // TrackPopupMenu needs special handling for proper message delivery
            DWORD cmd = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, 
                                       pt.x, pt.y, 0, hwnd, nullptr);
            
            // Post a fake message to ensure menu closes properly
            PostMessage(hwnd, WM_NULL, 0, 0);
            
            if (cmd == 100) {
                mgr->ClearClipboardHistory();
            } else if (cmd == 101 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->DeleteItem(clickedItemIndex);
            } else if (cmd == 102 && clickedItemIndex >= 0 && clickedItemIndex < listSize && mgr->snippetsMode) {
                mgr->PasteSnippet(mgr->filteredSnippetIndices[clickedItemIndex]);
                mgr->HideListWindow();
            } else if (cmd == 106 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                int savedSel = mgr->selectedIndex;
                mgr->selectedIndex = clickedItemIndex;
                bool ok = mgr->PasteToFocusedControlWithoutClipboard(false);
                mgr->selectedIndex = savedSel;
                if (!ok) {
                    MessageBoxW(hwnd, L"Could not paste: no text in history, or the target could not receive keystrokes.",
                                L"clip2", MB_OK | MB_ICONINFORMATION);
                } else {
                    mgr->HideListWindow();
                }
            } else if (cmd == 103 && mgr->snippetsMode) {
                mgr->ShowSnippetsManagerDialog();
            } else if (cmd == 104 && mgr->snippetsMode) {
                mgr->ShowSnippetsManagerDialog();
            } else if (cmd == 105 && clickedItemIndex >= 0 && clickedItemIndex < listSize && mgr->snippetsMode) {
                int actualIdx = mgr->filteredSnippetIndices[clickedItemIndex];
                if (actualIdx >= 0 && actualIdx < (int)mgr->snippets.size()) {
                    mgr->snippets.erase(mgr->snippets.begin() + actualIdx);
                    mgr->SaveSnippets();
                    mgr->FilterSnippets();
                    mgr->UpdateListWindow();
                }
            } else if (cmd == 110 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->TogglePin(clickedItemIndex);
            } else if (cmd == 111 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->CopyItemAsPlainText(clickedItemIndex);
            } else if (cmd == 112 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->OpenItemUrl(clickedItemIndex);
                mgr->HideListWindow();
            } else if (cmd == 113 && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->SaveImageItemToFile(clickedItemIndex);
            } else if (cmd >= TRANSFORM_UPPERCASE && cmd <= TRANSFORM_PLAIN_TEXT && clickedItemIndex >= 0 && clickedItemIndex < listSize && !mgr->snippetsMode) {
                mgr->TransformTextItem(clickedItemIndex, cmd);
            }
            
            DestroyMenu(hMenu);
            return 0;
        }
        
    case WM_MOUSEWHEEL:
        {
            int listSize = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            int maxScroll = std::max(0, listSize - mgr->itemsPerPage);
            int oldOffset = mgr->scrollOffset;
            
            if (delta > 0) {
                mgr->scrollOffset = std::max(0, mgr->scrollOffset - 3);
            } else {
                mgr->scrollOffset = std::min(maxScroll, mgr->scrollOffset + 3);
            }
            
            if (mgr->scrollOffset != oldOffset) {
                mgr->UpdateListWindow();
            }
            return 0;
        }
        
    case WM_SIZE:
        {
            // Update rounded region when window is resized
            RECT rect;
            GetClientRect(hwnd, &rect);
            HRGN hRgn = CreateRectRgn(0, 0, rect.right + 1, rect.bottom + 1);
            SetWindowRgn(hwnd, hRgn, TRUE);
            DeleteObject(hRgn);
            break;
        }
    
    case WM_CTLCOLOREDIT:
        {
            // Style the search edit control. Theme can change at runtime, so the cached
            // brush must be recreated when BG changes; otherwise we'd leak the old brush
            // (or worse, return one tied to a different theme).
            HDC hdcEdit = (HDC)wParam;
            SetTextColor(hdcEdit, Theme5250::TXT);
            SetBkColor(hdcEdit, Theme5250::BG);
            static HBRUSH hDarkBrush = nullptr;
            static COLORREF hDarkBrushColor = (COLORREF)~0u;
            if (!hDarkBrush || hDarkBrushColor != Theme5250::BG) {
                if (hDarkBrush) DeleteObject(hDarkBrush);
                hDarkBrush = CreateSolidBrush(Theme5250::BG);
                hDarkBrushColor = Theme5250::BG;
            }
            return (LRESULT)hDarkBrush;
        }
    
    case WM_KILLFOCUS:
        {
            // Hide list when it loses focus (unless focus is moving to a child or the preview popup)
            HWND hwndGettingFocus = (HWND)wParam;
            
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            bool isPreviewWindow = mgr->hwndPreview && hwndGettingFocus == mgr->hwndPreview;
            
            if (!isChildWindow && !isPreviewWindow && hwndGettingFocus != hwnd) {
                mgr->HideListWindow();
            } else {
                SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            }
        }
        break;
    
    case WM_ACTIVATE:
        if (wParam == WA_INACTIVE) {
            HWND hwndGettingFocus = GetFocus();
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            bool isPreviewWindow = mgr->hwndPreview && hwndGettingFocus == mgr->hwndPreview;
            
            if (!isChildWindow && !isPreviewWindow) {
                mgr->HideListWindow();
            }
        } else {
            SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        }
        break;
    
    case WM_NCACTIVATE:
        if (wParam == FALSE) {
            HWND hwndGettingFocus = GetFocus();
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            bool isPreviewWindow = mgr->hwndPreview && hwndGettingFocus == mgr->hwndPreview;
            
            if (!isChildWindow && !isPreviewWindow) {
                mgr->HideListWindow();
            }
        }
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK ClipboardManager::SearchEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    
    if (!mgr) {
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    
    // Handle Escape - close overlay from search box
    if (uMsg == WM_KEYDOWN && wParam == VK_ESCAPE) {
        mgr->HideListWindow();
        return 0;
    }
    
    // Handle Ctrl+F in search box - select all to start/refine search (works for both clipboard and snippets)
    if (uMsg == WM_KEYDOWN && wParam == 'F' && (GetAsyncKeyState(VK_CONTROL) & 0x8000)) {
        SendMessage(hwnd, EM_SETSEL, 0, -1);
        return 0;
    }
    
    // Ctrl+Left / Ctrl+Right switch between Clipboard and Snippets (from search box too)
    if (uMsg == WM_KEYDOWN && (GetAsyncKeyState(VK_CONTROL) & 0x8000) && (wParam == VK_LEFT || wParam == VK_RIGHT)) {
        SetFocus(mgr->hwndList);
        SendMessage(mgr->hwndList, WM_KEYDOWN, wParam, lParam);
        return 0;
    }
    
    // In snippets mode, forward Up/Down to list so arrow keys move selection even when search has focus
    if (uMsg == WM_KEYDOWN && mgr->snippetsMode && (wParam == VK_UP || wParam == VK_DOWN)) {
        SetFocus(mgr->hwndList);
        SendMessage(mgr->hwndList, WM_KEYDOWN, wParam, lParam);
        return 0;
    }
    
    // Handle Enter key in search box
    if (uMsg == WM_KEYDOWN && wParam == VK_RETURN) {
        int length = GetWindowTextLength(hwnd);
        if (length > 0) {
            wchar_t* buffer = new wchar_t[length + 1];
            GetWindowText(hwnd, buffer, length + 1);
            mgr->searchText = buffer;
            delete[] buffer;
        } else {
            mgr->searchText.clear();
        }
        if (mgr->snippetsMode) {
            if (_wcsicmp(mgr->searchText.c_str(), L"*set") == 0) {
                mgr->HideListWindow();
                mgr->ShowSnippetsManagerDialog();
            } else {
                mgr->FilterSnippets();
                if (!mgr->filteredSnippetIndices.empty()) {
                    mgr->PasteSnippet(mgr->filteredSnippetIndices[0]);
                    mgr->HideListWindow();
                }
            }
        } else {
            mgr->FilterItems();
            if (!mgr->filteredIndices.empty()) {
                mgr->PasteItem(0);
                mgr->HideListWindow();
            }
        }
        return 0;
    }
    
    // Call original window procedure for all other messages
    if (mgr->originalSearchEditProc) {
        return CallWindowProc(mgr->originalSearchEditProc, hwnd, uMsg, wParam, lParam);
    }
    
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void ClipboardManager::CreateTrayIcon() {
    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hwndMain;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    
    // Load icon from resource
    HICON hIcon = LoadIcon(GetModuleHandle(nullptr), MAKEINTRESOURCE(1));
    if (!hIcon) {
        // Fallback: try loading from file
        hIcon = (HICON)LoadImage(
            GetModuleHandle(nullptr),
            L"ico2.ico",
            IMAGE_ICON,
            GetSystemMetrics(SM_CXSMICON),
            GetSystemMetrics(SM_CYSMICON),
            LR_LOADFROMFILE
        );
    }
    
    if (hIcon) {
        nid.hIcon = hIcon;
    }
    
    wcscpy_s(nid.szTip, 128, L"clip2");
    Shell_NotifyIcon(NIM_ADD, &nid);
    
    if (hIcon) {
        DestroyIcon(hIcon);
    }
}

void ClipboardManager::UpdateTrayIcon() {
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void ClipboardManager::RemoveTrayIcon() {
    Shell_NotifyIcon(NIM_DELETE, &nid);
}

void ClipboardManager::RegisterHotkey() {
    if (!hwndMain) return;
    RegisterHotKey(hwndMain, HOTKEY_ID_OVERLAY, hotkeyConfig.modifiers | MOD_NOREPEAT, hotkeyConfig.vkCode);
#ifdef HAVE_UIAUTOMATION
    RegisterHotKey(hwndMain, HOTKEY_ID_COPY_FOCUSED, copyFocusedHotkey.modifiers | MOD_NOREPEAT, copyFocusedHotkey.vkCode);
#endif
    RegisterHotKey(hwndMain, HOTKEY_ID_PASTE_FOCUSED, pasteFocusedHotkey.modifiers | MOD_NOREPEAT, pasteFocusedHotkey.vkCode);
    RegisterHotKey(hwndMain, HOTKEY_ID_PASTE_FOCUSED_CLIPBOARD, pasteClipboardHotkey.modifiers | MOD_NOREPEAT, pasteClipboardHotkey.vkCode);
}

void ClipboardManager::UnregisterHotkey() {
    if (hwndMain) {
        UnregisterHotKey(hwndMain, HOTKEY_ID_OVERLAY);
#ifdef HAVE_UIAUTOMATION
        UnregisterHotKey(hwndMain, HOTKEY_ID_COPY_FOCUSED);
#endif
        UnregisterHotKey(hwndMain, HOTKEY_ID_PASTE_FOCUSED);
        UnregisterHotKey(hwndMain, HOTKEY_ID_PASTE_FOCUSED_CLIPBOARD);
    }
}

void ClipboardManager::InstallKeyboardHook() {
    // Install low-level keyboard hook for reliable NumPad key detection
    if (hKeyboardHook == nullptr) {
        hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandle(nullptr), 0);
    }
}

void ClipboardManager::UninstallKeyboardHook() {
    if (hKeyboardHook != nullptr) {
        UnhookWindowsHookEx(hKeyboardHook);
        hKeyboardHook = nullptr;
    }
}

LRESULT CALLBACK ClipboardManager::LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    
    if (nCode >= HC_ACTION && mgr != nullptr && mgr->hwndMain != nullptr) {
        KBDLLHOOKSTRUCT* pKeyboard = (KBDLLHOOKSTRUCT*)lParam;
        // Do not handle keys we inject ourselves (e.g. Ctrl+V from SendCtrlV) — breaks paste in target apps.
        if (pKeyboard->flags & (LLKHF_INJECTED | LLKHF_LOWER_IL_INJECTED))
            return CallNextHookEx(mgr->hKeyboardHook, nCode, wParam, lParam);

        // Check for configured hotkey modifiers
        bool hasCtrl = (mgr->hotkeyConfig.modifiers & MOD_CONTROL) != 0;
        bool hasAlt = (mgr->hotkeyConfig.modifiers & MOD_ALT) != 0;
        bool hasShift = (mgr->hotkeyConfig.modifiers & MOD_SHIFT) != 0;
        bool hasWin = (mgr->hotkeyConfig.modifiers & MOD_WIN) != 0;
        
        bool isCtrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        bool isAltPressed = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
        bool isShiftPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
        bool isWinPressed = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
        
        {
            // Snippets overlay shortcut is optional — if the user has not bound one
            // (vkCode == 0) we skip dispatch entirely so the keypress falls through.
            if (mgr->snippetsHotkey.vkCode != 0) {
                bool snipCtrl  = (mgr->snippetsHotkey.modifiers & MOD_CONTROL) != 0;
                bool snipAlt   = (mgr->snippetsHotkey.modifiers & MOD_ALT) != 0;
                bool snipShift = (mgr->snippetsHotkey.modifiers & MOD_SHIFT) != 0;
                bool snipWin   = (mgr->snippetsHotkey.modifiers & MOD_WIN) != 0;
                bool snipMatch = (snipCtrl == isCtrlPressed) && (snipAlt == isAltPressed) &&
                                 (snipShift == isShiftPressed) && (snipWin == isWinPressed);
                if (wParam == WM_KEYDOWN && snipMatch && pKeyboard->vkCode == mgr->snippetsHotkey.vkCode) {
                    PostMessage(mgr->hwndMain, ClipboardManager::WM_SNIPPETS_OVERLAY_HOTKEY, 0, 0);
                    return 1;
                }
            }
        }
        
        // Plain Esc closes overlay when list, search, preview, or a list child has focus
        if (wParam == WM_KEYDOWN && pKeyboard->vkCode == VK_ESCAPE && mgr->listVisible) {
            bool plainEsc = !isCtrlPressed && !isAltPressed && !isShiftPressed && !isWinPressed;
            if (plainEsc) {
                HWND fg = GetForegroundWindow();
                if (fg == mgr->hwndList || fg == mgr->hwndSearch ||
                    (mgr->hwndPreview && fg == mgr->hwndPreview) ||
                    (fg && mgr->hwndList && IsChild(mgr->hwndList, fg))) {
                    PostMessage(mgr->hwndMain, ClipboardManager::WM_DISMISS_OVERLAY, 0, 0);
                    return 1;
                }
            }
        }
        
        // Check if all required modifiers are pressed
        bool modifiersMatch = true;
        if (hasCtrl && !isCtrlPressed) modifiersMatch = false;
        if (hasAlt && !isAltPressed) modifiersMatch = false;
        if (hasShift && !isShiftPressed) modifiersMatch = false;
        if (hasWin && !isWinPressed) modifiersMatch = false;
        
        // Also check that unwanted modifiers are NOT pressed (strict matching)
        if ((!hasCtrl && isCtrlPressed) || (!hasAlt && isAltPressed) || 
            (!hasShift && isShiftPressed) || (!hasWin && isWinPressed)) {
            modifiersMatch = false;
        }
        
        // Check if the configured key is pressed
        if (wParam == WM_KEYDOWN && modifiersMatch && pKeyboard->vkCode == mgr->hotkeyConfig.vkCode) {
            // Post message to main window to toggle list
            PostMessage(mgr->hwndMain, WM_CLIPBOARD_HOTKEY, 0, 0);
            return 1; // Consume the key so it doesn't propagate
        }
    }
    
    HHOOK hook = mgr ? mgr->hKeyboardHook : nullptr;
    return CallNextHookEx(hook, nCode, wParam, lParam);
}

void ClipboardManager::SetListSnippetsMode(bool wantSnippets) {
    if (snippetsMode == wantSnippets) {
        return;
    }
    snippetsMode = wantSnippets;
    numberInput.clear();
    scrollOffset = 0;
    if (snippetsMode) {
        searchText.clear();
        if (hwndSearch) {
            SetWindowText(hwndSearch, L"");
            SendMessage(hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search snippets (Ctrl+F) | Arrow keys to move, Enter to paste");
        }
        FilterSnippets();
        selectedIndex = (filteredSnippetIndices.empty() ? -1 : 0);
        SetFocus(hwndList);
    } else {
        FilterItems();
        if (hwndSearch) {
            SendMessage(hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search... (Ctrl+F)");
        }
    }
    UpdateListWindow();
}

void ClipboardManager::ShowListWindow(bool startInSnippetsMode) {
    if (!hwndList) {
        return;
    }
    
    // Make sure any existing timer is stopped
    KillTimer(hwndMain, 1);
    
    // Remember the window that had focus before we show our list
    previousFocusWindow = GetForegroundWindow();
    
    // Reset scroll to top when showing
    scrollOffset = 0;
    // Clear number input when showing
    numberInput.clear();
    // Clear search text when showing
    searchText.clear();
    if (hwndSearch) {
        SetWindowText(hwndSearch, L"");
    }
    if (startInSnippetsMode) {
        snippetsMode = true;
        if (hwndSearch) {
            SendMessage(hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search snippets (Ctrl+F) | Arrow keys to move, Enter to paste");
        }
        FilterSnippets();
        if (!filteredSnippetIndices.empty()) {
            selectedIndex = 0;
        } else {
            selectedIndex = -1;
        }
    } else {
        snippetsMode = false;
        FilterItems();
        if (!filteredIndices.empty()) {
            selectedIndex = 0;
        } else {
            selectedIndex = -1;
        }
    }
    ClearMultiSelection(); // Clear multi-selection when showing list
    // Show search box
    if (hwndSearch) {
        ShowWindow(hwndSearch, SW_SHOW);
    }

    UpdateListWindow();
    
    // Position window at center-top of screen
    RECT screenRect;
    GetWindowRect(GetDesktopWindow(), &screenRect);
    int x = (screenRect.right - WINDOW_WIDTH) / 2;
    int y = screenRect.top + 50;
    SetWindowPos(hwndList, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
    
    ShowWindow(hwndList, SW_SHOW);
    FocusListWindow();
    // Make sure list window has focus (not search box) so arrow keys work immediately
    SetFocus(hwndList);
    listVisible = true;
    
    // Start timer to check if window loses focus (check every 100ms)
    SetTimer(hwndMain, 1, 100, nullptr);
}

void ClipboardManager::HideListWindow() {
    if (listVisible) {
        HidePreviewWindow();
        // Clear search text when hiding
        searchText.clear();
        if (hwndSearch) {
            SetWindowText(hwndSearch, L"");
            ShowWindow(hwndSearch, SW_HIDE);
        }
        ShowWindow(hwndList, SW_HIDE);
        listVisible = false;
        
        // Stop the focus check timer
        KillTimer(hwndMain, 1);
    }
}

int ClipboardManager::GetItemAtPosition(int x, int y) {
    if (!hwndList) return -1;
    
    RECT rect;
    GetClientRect(hwndList, &rect);
    
    int itemHeight = 50;
    int clickedY = y - 60; // content starts at CONTENT_TOP (60): header (28) + search (to 56) + gap
    
    if (clickedY < 0) return -1;
    
    int row = clickedY / itemHeight;
    if (itemsPerPage > 0 && row >= itemsPerPage) return -1;  // below the last drawn row (footer gap)
    int filteredIndex = scrollOffset + row;
    
    if (snippetsMode) {
        if (filteredIndex >= 0 && filteredIndex < (int)filteredSnippetIndices.size()) {
            return filteredIndex;
        }
    } else {
        if (filteredIndex >= 0 && filteredIndex < (int)filteredIndices.size()) {
            return filteredIndex;
        }
    }
    return -1;
}

void ClipboardManager::ShowPreviewWindow(int itemIndex, int mouseX, int mouseY) {
    // itemIndex is a filtered index, convert to actual index
    if (!hwndPreview || itemIndex < 0 || itemIndex >= (int)filteredIndices.size()) {
        return;
    }
    
    int actualIndex = filteredIndices[itemIndex];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) {
        return;
    }
    
    const auto& item = clipboardHistory[actualIndex];
    if (!item->isImage && !item->isVideo) {
        return;
    }
    
    // Calculate preview size based on image
    // Base size matches the window creation size
    const int BASE_PREVIEW_SIZE = 400;
    int previewWidth = BASE_PREVIEW_SIZE;
    int previewHeight = BASE_PREVIEW_SIZE;
    
    HBITMAP previewBmp = item->GetPreviewBitmap();
    if (previewBmp) {
        BITMAP bm;
        GetObject(previewBmp, sizeof(BITMAP), &bm);
        // Use larger max size for better preview quality
        int maxSize = BASE_PREVIEW_SIZE - 20; // Account for padding
        float scale = std::min((float)maxSize / bm.bmWidth, (float)maxSize / bm.bmHeight);
        previewWidth = (int)(bm.bmWidth * scale) + 20; // Add padding for border
        previewHeight = (int)(bm.bmHeight * scale) + 20;
        // Cap at reasonable max to prevent huge windows
        previewWidth = std::min(previewWidth, 600);
        previewHeight = std::min(previewHeight, 600);
        // Ensure minimum size
        previewWidth = std::max(previewWidth, 200);
        previewHeight = std::max(previewHeight, 200);
    }
    
    // Position preview window near mouse cursor (offset to avoid covering cursor)
    RECT screenRect;
    GetWindowRect(GetDesktopWindow(), &screenRect);
    
    int previewX = mouseX + 20;
    int previewY = mouseY + 20;
    
    // Adjust if would go off screen
    if (previewX + previewWidth > screenRect.right) {
        previewX = mouseX - previewWidth - 20;
    }
    if (previewY + previewHeight > screenRect.bottom) {
        previewY = mouseY - previewHeight - 20;
    }
    
    SetWindowPos(hwndPreview, HWND_TOPMOST, previewX, previewY, previewWidth, previewHeight, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    
    // Update rounded region to match new window size
    HRGN hPreviewRgn = CreateRectRgn(0, 0, previewWidth + 1, previewHeight + 1);
    SetWindowRgn(hwndPreview, hPreviewRgn, TRUE);
    DeleteObject(hPreviewRgn);
    
    InvalidateRect(hwndPreview, nullptr, TRUE);
    UpdateWindow(hwndPreview);
}

void ClipboardManager::HidePreviewWindow() {
    if (hwndPreview) {
        ShowWindow(hwndPreview, SW_HIDE);
        hoveredItemIndex = -1;
    }
}

void ClipboardManager::FocusListWindow() {
    if (!hwndList) {
        return;
    }

    HWND fgWindow = GetForegroundWindow();
    DWORD currentThread = GetCurrentThreadId();
    DWORD fgThread = fgWindow ? GetWindowThreadProcessId(fgWindow, nullptr) : 0;

    if (fgThread != 0 && fgThread != currentThread) {
        AttachThreadInput(fgThread, currentThread, TRUE);
    }

    SetWindowPos(hwndList, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    SetForegroundWindow(hwndList);
    SetActiveWindow(hwndList);
    SetFocus(hwndList);
    BringWindowToTop(hwndList);

    if (fgThread != 0 && fgThread != currentThread) {
        AttachThreadInput(fgThread, currentThread, FALSE);
    }
}

void ClipboardManager::UpdateListWindow() {
    if (hwndList) {
        InvalidateRect(hwndList, nullptr, TRUE);
        UpdateWindow(hwndList);
    }
}

void ClipboardManager::PasteItem(int index) {
    if (index < 0 || index >= (int)filteredIndices.size()) return;
    int actualIndex = filteredIndices[index];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) return;

    const auto& item = clipboardHistory[actualIndex];
    if (!item) return;

    bool pasteAsPlainText = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;

    isPasting = true;

    std::wstring plainText = GetPlainTextForDirectPaste(item.get());

    if (!OpenClipboardWithRetry(hwndMain, 24, 8)) {
        isPasting = false;
        return;
    }
    EmptyClipboard();

    bool success = false;
    if (pasteAsPlainText) {
        std::wstring t = plainText.empty() ? item->preview : plainText;
        if (!t.empty()) {
            size_t bytes = (t.size() + 1) * sizeof(wchar_t);
            success = PutBytesOnClipboard(CF_UNICODETEXT, t.c_str(), bytes);
        }
    } else if (!item->formats.empty()) {
        for (const auto& fp : item->formats) {
            if (fp.second.empty()) continue;
            if (PutBytesOnClipboard(fp.first, fp.second.data(), fp.second.size()))
                success = true;
        }
        if (!success && !plainText.empty()) {
            size_t bytes = (plainText.size() + 1) * sizeof(wchar_t);
            success = PutBytesOnClipboard(CF_UNICODETEXT, plainText.c_str(), bytes);
        }
    } else if (!plainText.empty()) {
        size_t bytes = (plainText.size() + 1) * sizeof(wchar_t);
        success = PutBytesOnClipboard(CF_UNICODETEXT, plainText.c_str(), bytes);
    }
    CloseClipboard();

    if (!success) {
        isPasting = false;
        return;
    }

    if (!plainText.empty())
        lastPastedText = plainText;
    else
        lastPastedText = item->preview;
    lastSequenceNumber = GetClipboardSequenceNumber();

    if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
        SetForegroundWindow(previousFocusWindow);
        SetFocus(previousFocusWindow);
        Sleep(40);
    }
    ReleaseHotkeyModifiersForPaste();
    Sleep(20);

    SendCtrlV();

    // Keep isPasting set for a while: some apps echo what was pasted right back to the clipboard,
    // and we want the duplicate detector to ignore that.
    Sleep(420);
    isPasting = false;
}

void ClipboardManager::PasteMultipleItems() {
    if (multiSelectedIndices.empty()) return;

    // Sort indices so pasting follows visible top-to-bottom order.
    std::vector<int> sortedIndices(multiSelectedIndices.begin(), multiSelectedIndices.end());
    std::sort(sortedIndices.begin(), sortedIndices.end());

    std::vector<const ClipboardItem*> selectedItems;
    selectedItems.reserve(sortedIndices.size());
    for (int filteredIndex : sortedIndices) {
        if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) continue;
        int actualIndex = filteredIndices[filteredIndex];
        if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) continue;
        selectedItems.push_back(clipboardHistory[actualIndex].get());
    }

    if (selectedItems.empty()) { ClearMultiSelection(); return; }

    // Ctrl+Enter forces plain text mode and skips RTF/HTML.
    bool pasteAsPlainText = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;

    // Decide fast vs slow path. Fast path: every selected item is text/RTF/HTML only — we
    // can combine them into ONE clipboard payload and paste once. Slow path: at least one
    // item carries an image or other binary clipboard format and must be replayed in turn.
    bool fastPath = true;
    for (const ClipboardItem* it : selectedItems) {
        if (!it) continue;
        if (!ItemHasOnlyTextFormats(it)) { fastPath = false; break; }
        std::wstring t = GetPlainTextForDirectPaste(it);
        if (t.empty()) { fastPath = false; break; }
    }

    isPasting = true;

    if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
        SetForegroundWindow(previousFocusWindow);
        SetFocus(previousFocusWindow);
        Sleep(70);
    }
    ReleaseHotkeyModifiersForPaste();

    std::map<UINT, std::vector<BYTE>> backup;
    if (!BackupClipboardSerialFormats(hwndMain, backup)) {
        isPasting = false;
        ClearMultiSelection();
        return;
    }

    if (fastPath) {
        // ---- FAST PATH: combine all selected items into one clipboard payload ----
        const wchar_t* sep = L"\r\n";
        std::wstring combined;
        bool anyRtfSource = false;
        UINT cfRtf = CfRtf();
        for (size_t i = 0; i < selectedItems.size(); i++) {
            const ClipboardItem* it = selectedItems[i];
            if (!it) continue;
            std::wstring t = GetPlainTextForDirectPaste(it);
            if (i > 0) combined += sep;
            combined += t;
            if (cfRtf != 0) {
                const std::vector<BYTE>* rtf = it->GetFormatData(cfRtf);
                if (rtf && !rtf->empty()) anyRtfSource = true;
            }
        }

        bool clipboardSet = false;
        if (OpenClipboardWithRetry(hwndMain, 30, 8)) {
            EmptyClipboard();
            if (!combined.empty()) {
                size_t bytes = (combined.size() + 1) * sizeof(wchar_t);
                if (PutBytesOnClipboard(CF_UNICODETEXT, combined.c_str(), bytes))
                    clipboardSet = true;
            }
            if (!pasteAsPlainText && anyRtfSource && cfRtf != 0 && !combined.empty()) {
                // Provide an RTF doc too so rich-text targets receive a real RTF stream.
                std::string rtf = BuildRtfWrapFromUnicode(combined);
                if (!rtf.empty())
                    PutBytesOnClipboard(cfRtf, rtf.data(), rtf.size() + 1);
            }
            CloseClipboard();
        }

        if (!clipboardSet) {
            RestoreClipboardSerialFormats(hwndMain, backup);
            isPasting = false;
            ClearMultiSelection();
            return;
        }

        lastPastedText = combined;
        lastSequenceNumber = GetClipboardSequenceNumber();
        Sleep(30);
        SendCtrlV();
        Sleep(320);
        RestoreClipboardSerialFormats(hwndMain, backup);
        lastSequenceNumber = GetClipboardSequenceNumber();
        Sleep(70);
        isPasting = false;
        ClearMultiSelection();
        PlayClickSound();
        return;
    }

    // ---- SLOW PATH: at least one non-text item; replay each item via swap+Ctrl+V ----
    // Longer pacing than before because the previous timings were too aggressive on
    // slower editors (Office apps, browsers) and dropped trailing items.
    const DWORD kFocusMs           = 70;
    const DWORD kAfterItemMs       = 200;
    const DWORD kAfterSeparatorMs  = 140;
    const DWORD kTailBeforeRestore = 200;
    (void)kFocusMs;

    auto waitForClipboardReady = [&]() {
        for (int attempt = 0; attempt < 32; attempt++) {
            if (OpenClipboard(hwndMain)) { CloseClipboard(); return true; }
            Sleep(8);
        }
        return false;
    };

    std::wstring pastedTextAggregate;

    for (size_t i = 0; i < selectedItems.size(); i++) {
        const ClipboardItem* item = selectedItems[i];
        if (!item) continue;

        std::wstring plainText = GetPlainTextForDirectPaste(item);
        waitForClipboardReady();

        bool clipboardSet = false;
        if (!pasteAsPlainText && !item->formats.empty())
            clipboardSet = SetClipboardFromHistoryItem(hwndMain, item);
        if (!clipboardSet && !plainText.empty())
            clipboardSet = SetClipboardUnicodeOnly(hwndMain, plainText);
        if (!clipboardSet) continue;

        if (!plainText.empty()) {
            if (!pastedTextAggregate.empty()) pastedTextAggregate += L"\n";
            pastedTextAggregate += plainText;
        }

        lastSequenceNumber = GetClipboardSequenceNumber();
        Sleep(20);
        SendCtrlV();
        Sleep(kAfterItemMs);

        if (i + 1 < selectedItems.size()) {
            waitForClipboardReady();
            if (SetClipboardUnicodeOnly(hwndMain, L"\r\n")) {
                lastSequenceNumber = GetClipboardSequenceNumber();
                Sleep(20);
                SendCtrlV();
                Sleep(kAfterSeparatorMs);
            }
        }
    }

    Sleep(kTailBeforeRestore);
    RestoreClipboardSerialFormats(hwndMain, backup);
    lastSequenceNumber = GetClipboardSequenceNumber();
    if (!pastedTextAggregate.empty())
        lastPastedText = pastedTextAggregate;

    Sleep(80);
    isPasting = false;
    ClearMultiSelection();
    PlayClickSound();
}

void ClipboardManager::ToggleMultiSelect(int filteredIndex) {
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) {
        return;
    }
    
    // Toggle selection
    if (multiSelectedIndices.find(filteredIndex) != multiSelectedIndices.end()) {
        multiSelectedIndices.erase(filteredIndex);
    } else {
        multiSelectedIndices.insert(filteredIndex);
    }
    
    UpdateListWindow();
}

void ClipboardManager::ClearMultiSelection() {
    multiSelectedIndices.clear();
    multiSelectAnchor = -1;
    UpdateListWindow();
}

// Try to grab clipboard content immediately (e.g. before a blocking app clears it). Call from WM_CLIPBOARDUPDATE.
bool ClipboardManager::TryCaptureClipboardImmediately() {
    hasImmediateClipboardSnapshot = false;
    immediateClipboardSnapshot.clear();
    if (!hwndMain) return false;
    if (!OpenClipboard(hwndMain)) return false;

    const UINT priorityFormats[] = {
        CF_HDROP, CF_UNICODETEXT, CF_TEXT, CF_BITMAP, CF_DIBV5, CF_DIB,
        CF_ENHMETAFILE, CF_METAFILEPICT
    };
    auto isHandleBased = [](UINT fmt) {
        return fmt == CF_BITMAP || fmt == CF_PALETTE || fmt == CF_METAFILEPICT ||
               fmt == CF_ENHMETAFILE || fmt == 0x0082 || fmt == 0x008E || fmt == 0x0083;
    };

    try {
        UINT primaryFormat = 0;
        for (UINT pf : priorityFormats) {
            if (IsClipboardFormatAvailable(pf)) { primaryFormat = pf; break; }
        }
        if (primaryFormat == 0) primaryFormat = EnumClipboardFormats(0);
        if (primaryFormat == 0) { CloseClipboard(); return false; }

        auto copyFormat = [&isHandleBased](UINT fmt, std::vector<BYTE>& out) -> bool {
            if (isHandleBased(fmt) && fmt != CF_BITMAP) return false;
            if (fmt == CF_BITMAP) {
                HBITMAP hBmp = (HBITMAP)GetClipboardData(CF_BITMAP);
                if (!hBmp) return false;
                out = ConvertBitmapToDIB(hBmp);
                return !out.empty();
            }
            HGLOBAL hMem = GetClipboardData(fmt);
            if (!hMem) return false;
            SIZE_T sz = GlobalSize(hMem);
            if (sz == 0 || sz >= 100 * 1024 * 1024) return false;
            void* p = GlobalLock(hMem);
            if (!p) return false;
            out.resize(sz);
            memcpy(out.data(), p, sz);
            GlobalUnlock(hMem);
            return true;
        };

        std::vector<BYTE> primaryData;
        UINT storageFormat = primaryFormat;
        if (primaryFormat == CF_BITMAP) {
            if (!copyFormat(CF_BITMAP, primaryData)) { CloseClipboard(); return false; }
            storageFormat = CF_DIB;
        } else if (isHandleBased(primaryFormat)) {
            CloseClipboard();
            return false;
        } else {
            if (!copyFormat(primaryFormat, primaryData)) { CloseClipboard(); return false; }
        }
        if (primaryData.empty() || primaryData.size() >= 200 * 1024 * 1024) {
            CloseClipboard();
            return false;
        }
        immediateClipboardSnapshot[storageFormat] = std::move(primaryData);

        const int MAX_FORMATS_PER_ITEM = 12;
        const size_t MAX_FORMAT_SIZE_PER_ITEM = 10 * 1024 * 1024;
        int formatCount = 1;
        size_t totalSize = immediateClipboardSnapshot[storageFormat].size();
        UINT format = EnumClipboardFormats(0);
        while (format != 0 && formatCount < MAX_FORMATS_PER_ITEM) {
            if (format == primaryFormat || format == storageFormat) {
                format = EnumClipboardFormats(format);
                continue;
            }
            if (isHandleBased(format)) {
                format = EnumClipboardFormats(format);
                continue;
            }
            std::vector<BYTE> data;
            if (!copyFormat(format, data) || data.empty() || data.size() >= 5 * 1024 * 1024 ||
                totalSize + data.size() > MAX_FORMAT_SIZE_PER_ITEM) {
                format = EnumClipboardFormats(format);
                continue;
            }
            immediateClipboardSnapshot[format] = std::move(data);
            totalSize += immediateClipboardSnapshot[format].size();
            formatCount++;
            format = EnumClipboardFormats(format);
        }
        hasImmediateClipboardSnapshot = true;
    } catch (...) {
        immediateClipboardSnapshot.clear();
    }
    CloseClipboard();
    return hasImmediateClipboardSnapshot;
}

// Process a previously captured snapshot (bypasses apps that clear the clipboard after copy).
void ClipboardManager::ProcessClipboardFromSnapshot() {
    if (immediateClipboardSnapshot.empty()) return;
    const UINT priorityFormats[] = {
        CF_HDROP, CF_UNICODETEXT, CF_TEXT, CF_DIB, CF_DIBV5, CF_ENHMETAFILE, CF_METAFILEPICT
    };
    UINT primaryFormat = 0;
    for (UINT pf : priorityFormats) {
        auto it = immediateClipboardSnapshot.find(pf);
        if (it != immediateClipboardSnapshot.end() && !it->second.empty()) {
            primaryFormat = pf;
            break;
        }
    }
    if (primaryFormat == 0 && !immediateClipboardSnapshot.empty())
        primaryFormat = immediateClipboardSnapshot.begin()->first;
    if (primaryFormat == 0) {
        immediateClipboardSnapshot.clear();
        hasImmediateClipboardSnapshot = false;
        return;
    }
    const std::vector<BYTE>& primaryData = immediateClipboardSnapshot[primaryFormat];
    if (primaryData.empty() || primaryData.size() >= 200 * 1024 * 1024) {
        immediateClipboardSnapshot.clear();
        hasImmediateClipboardSnapshot = false;
        return;
    }
    std::unique_ptr<ClipboardItem> item;
    try {
        item = std::make_unique<ClipboardItem>(primaryFormat, primaryData);
    } catch (...) {
        immediateClipboardSnapshot.clear();
        hasImmediateClipboardSnapshot = false;
        return;
    }
    for (const auto& kv : immediateClipboardSnapshot) {
        if (kv.first == primaryFormat || kv.second.empty()) continue;
        item->AddFormat(kv.first, kv.second);
    }
    bool isPastPaste = false;
    if (!lastPastedText.empty() && item->format == CF_UNICODETEXT) {
        try {
            const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
            if (textData && textData->size() >= sizeof(wchar_t)) {
                size_t len = textData->size() / sizeof(wchar_t);
                if (len > 0) {
                    const wchar_t* textPtr = (const wchar_t*)textData->data();
                    size_t actualLen = len;
                    for (size_t j = 0; j < len; j++) {
                        if (textPtr[j] == L'\0') { actualLen = j; break; }
                    }
                    if (actualLen > 0) {
                        std::wstring clipboardText(textPtr, actualLen);
                        if (clipboardText == lastPastedText) {
                            isPastPaste = true;
                            lastPastedText.clear();
                        }
                    }
                }
            }
        } catch (...) {}
    }
    bool isDuplicate = false;
    if (!clipboardHistory.empty() && item && item->formats.size() > 0 && clipboardHistory[0]->formats.size() > 0) {
        try {
            const auto& lastItem = clipboardHistory[0];
            if (item->formats.size() == lastItem->formats.size()) {
                bool allMatch = true;
                const size_t MAX_COMPARE = 10 * 1024 * 1024;
                for (const auto& fp : item->formats) {
                    const std::vector<BYTE>* lastData = lastItem->GetFormatData(fp.first);
                    if (!lastData || lastData->size() != fp.second.size() || fp.second.size() > MAX_COMPARE ||
                        (fp.second.size() > 0 && memcmp(fp.second.data(), lastData->data(), fp.second.size()) != 0)) {
                        allMatch = false;
                        break;
                    }
                }
                if (allMatch) isDuplicate = true;
            }
        } catch (...) {}
    }
    // Same text copied twice (e.g. Ctrl+C twice on "text") → treat as duplicate even if raw bytes differ
    if (!isDuplicate && !clipboardHistory.empty() && item) {
        std::wstring newNorm = GetNormalizedTextForDuplicateCheck(item.get());
        std::wstring lastNorm = GetNormalizedTextForDuplicateCheck(clipboardHistory[0].get());
        if (!newNorm.empty() && newNorm == lastNorm)
            isDuplicate = true;
    }
    if (!isDuplicate && !isPastPaste && item) {
        try {
            clipboardHistory.insert(clipboardHistory.begin(), std::move(item));
            TrimHistory();
            if (listVisible) { FilterItems(); UpdateListWindow(); }
        } catch (...) {}
    }
    immediateClipboardSnapshot.clear();
    hasImmediateClipboardSnapshot = false;
}

#ifdef HAVE_UIAUTOMATION
// BSTR may contain embedded WChars; std::wstring(bstr) stops at first L'\0'.
static std::wstring UiaBstrToWstring(BSTR bstr) {
    if (!bstr) return L"";
    UINT n = SysStringLen(bstr);
    return std::wstring(bstr, (size_t)n);
}

// Returns true if s is a known placeholder / "not accessible" message (not real content).
static bool IsPlaceholderOrNotAccessibleText(const std::wstring& s) {
    if (s.empty()) return true;
    std::wstring lower;
    lower.reserve(s.size());
    for (wchar_t c : s) lower += (wchar_t)towlower(c);
    // "The editor is not accessible at this time. To enable screen reader optimized mode, use Shift+Alt+F1"
    if (lower.find(L"not accessible at this time") != std::wstring::npos) return true;
    if (lower.find(L"screen reader optimized mode") != std::wstring::npos) return true;
    return false;
}

// Get text from a UI Automation element (selection or document range). Returns true if text was written.
static bool GetTextFromElement(IUIAutomation* pAutomation, IUIAutomationElement* pElement, std::wstring& outText) {
    if (!pAutomation || !pElement) return false;
    IUnknown* pPatternUnknown = nullptr;
    HRESULT hr = pElement->GetCurrentPattern(UIA_TextPatternId, (IUnknown**)&pPatternUnknown);
    if (FAILED(hr) || !pPatternUnknown) return false;
    IUIAutomationTextPattern* pTextPattern = nullptr;
    hr = pPatternUnknown->QueryInterface(__uuidof(IUIAutomationTextPattern), (void**)&pTextPattern);
    pPatternUnknown->Release();
    if (FAILED(hr) || !pTextPattern) return false;
    std::wstring text;
    IUIAutomationTextRangeArray* pRangeArray = nullptr;
    hr = pTextPattern->GetSelection(&pRangeArray);
    if (SUCCEEDED(hr) && pRangeArray) {
        int length = 0;
        if (SUCCEEDED(pRangeArray->get_Length(&length)) && length > 0) {
            IUIAutomationTextRange* pRange = nullptr;
            if (SUCCEEDED(pRangeArray->GetElement(0, &pRange)) && pRange) {
                BSTR bstr = nullptr;
                if (SUCCEEDED(pRange->GetText(-1, &bstr)) && bstr) {
                    text = UiaBstrToWstring(bstr);
                    SysFreeString(bstr);
                }
                pRange->Release();
            }
        }
        pRangeArray->Release();
    }
    if (text.empty()) {
        IUIAutomationTextRange* pDocRange = nullptr;
        hr = pTextPattern->get_DocumentRange(&pDocRange);
        if (SUCCEEDED(hr) && pDocRange) {
            BSTR bstr = nullptr;
            if (SUCCEEDED(pDocRange->GetText(-1, &bstr)) && bstr) {
                text = UiaBstrToWstring(bstr);
                SysFreeString(bstr);
            }
            pDocRange->Release();
        }
    }
    pTextPattern->Release();
    if (text.empty()) return false;
    outText = std::move(text);
    return true;
}
#endif

// Get text from the currently focused control via UI Automation (bypasses apps that never put content on clipboard).
bool ClipboardManager::CopyFromFocusedControlViaUIA() {
#ifdef HAVE_UIAUTOMATION
    IUIAutomation* pAutomation = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER,
                                  __uuidof(IUIAutomation), (void**)&pAutomation);
    if (FAILED(hr) || !pAutomation) return false;

    std::wstring text;
    IUIAutomationElement* pFocused = nullptr;
    hr = pAutomation->GetFocusedElement(&pFocused);
    if (SUCCEEDED(hr) && pFocused) {
        GetTextFromElement(pAutomation, pFocused, text);
        pFocused->Release();
    }
    // If focused element gave a placeholder (e.g. "The editor is not accessible at this time..."), try foreground window
    if (IsPlaceholderOrNotAccessibleText(text)) {
        text.clear();
        HWND hFore = GetForegroundWindow();
        if (hFore) {
            IUIAutomationElement* pWindow = nullptr;
            hr = pAutomation->ElementFromHandle(hFore, &pWindow);
            if (SUCCEEDED(hr) && pWindow) {
                GetTextFromElement(pAutomation, pWindow, text);
                pWindow->Release();
            }
        }
    }
    pAutomation->Release();

    if (text.empty() || IsPlaceholderOrNotAccessibleText(text)) return false;

    // Add to our history and set system clipboard so user can paste
    std::vector<BYTE> data((text.size() + 1) * sizeof(wchar_t));
    memcpy(data.data(), text.c_str(), (text.size() + 1) * sizeof(wchar_t));
    std::unique_ptr<ClipboardItem> item;
    try {
        item = std::make_unique<ClipboardItem>(CF_UNICODETEXT, data);
    } catch (...) {
        return true;  // We got text; clipboard set below may still work
    }
    clipboardHistory.insert(clipboardHistory.begin(), std::move(item));
    TrimHistory();
    if (listVisible) { FilterItems(); UpdateListWindow(); }

    if (hwndMain && OpenClipboard(hwndMain)) {
        EmptyClipboard();
        size_t byteLen = (text.size() + 1) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, byteLen);
        if (hMem) {
            void* p = GlobalLock(hMem);
            if (p) {
                memcpy(p, text.c_str(), byteLen);
                GlobalUnlock(hMem);
                SetClipboardData(CF_UNICODETEXT, hMem);
                lastSequenceNumber = GetClipboardSequenceNumber();
            } else {
                GlobalFree(hMem);
            }
        }
        CloseClipboard();
    }
    PlayClickSound();
    return true;
#else
    (void)0;
    return false;
#endif
}

bool ClipboardManager::PasteToFocusedControlWithoutClipboard(bool useClipboardSwap) {
    if (clipboardHistory.empty())
        return false;
    int actualIndex = -1;
    if (listVisible && selectedIndex >= 0 && selectedIndex < (int)filteredIndices.size())
        actualIndex = filteredIndices[selectedIndex];
    else if (!clipboardHistory.empty())
        actualIndex = 0;
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size())
        return false;
    const ClipboardItem* item = clipboardHistory[actualIndex].get();
    std::wstring text = GetPlainTextForDirectPaste(item);
    if (text.empty() && item->formats.empty())
        return false;
    HWND hTarget = GetForegroundWindow();
    if (hTarget == hwndList || hTarget == hwndMain || (hwndList && IsChild(hwndList, hTarget))) {
        if (previousFocusWindow && IsWindow(previousFocusWindow) && previousFocusWindow != hwndList && previousFocusWindow != hwndMain)
            hTarget = previousFocusWindow;
        else
            hTarget = nullptr;
    }
    if (!hTarget || !IsWindow(hTarget))
        return false;

    DWORD tgtTid = GetWindowThreadProcessId(hTarget, nullptr);
    DWORD curTid = GetCurrentThreadId();
    if (tgtTid != curTid)
        AttachThreadInput(curTid, tgtTid, TRUE);
    SetForegroundWindow(hTarget);
    if (tgtTid != curTid)
        AttachThreadInput(curTid, tgtTid, FALSE);

    Sleep(80);
    ReleaseHotkeyModifiersForPaste();

    isPasting = true;

    if (!useClipboardSwap) {
        if (text.empty()) {
            isPasting = false;
            return false;
        }
        // SendUnicodeTextAsKeystrokes handles batching, partial-send retries, and \r\n.
        bool ok = SendUnicodeTextAsKeystrokes(text);
        if (ok) {
            lastPastedText = text;
            // Give the receiving app a beat before the next event so trailing chars actually land.
            Sleep(60);
        }
        isPasting = false;
        if (ok) PlayClickSound();
        return ok;
    }

    std::map<UINT, std::vector<BYTE>> backup;
    if (!BackupClipboardSerialFormats(hwndMain, backup)) {
        isPasting = false;
        return false;
    }

    bool clipboardSet = false;
    if (!item->formats.empty())
        clipboardSet = SetClipboardFromHistoryItem(hwndMain, item);
    if (!clipboardSet && !text.empty())
        clipboardSet = SetClipboardUnicodeOnly(hwndMain, text);
    if (!clipboardSet) {
        RestoreClipboardSerialFormats(hwndMain, backup);
        isPasting = false;
        return false;
    }

    if (!text.empty())
        lastPastedText = text;
    else
        lastPastedText = LastPastedTextFromItem(item);

    lastSequenceNumber = GetClipboardSequenceNumber();

    Sleep(20);
    SendCtrlV();
    Sleep(260);
    RestoreClipboardSerialFormats(hwndMain, backup);
    lastSequenceNumber = GetClipboardSequenceNumber();

    isPasting = false;
    PlayClickSound();
    return true;
}

void ClipboardManager::ProcessClipboard() {
    // Prevent re-entrant calls
    if (isProcessingClipboard) {
        return;
    }
    
    isProcessingClipboard = true;
    
    // If we captured the clipboard immediately (to bypass copy blocks), process that snapshot and exit
    if (hasImmediateClipboardSnapshot) {
        ProcessClipboardFromSnapshot();
        isProcessingClipboard = false;
        return;
    }
    
    // Retry logic with delays to avoid interfering with copy operations
    const int MAX_RETRIES = 5;
    const int RETRY_DELAY_MS = 50;
    
    for (int retry = 0; retry < MAX_RETRIES; retry++) {
        // Wait before retrying (except first attempt)
        if (retry > 0) {
            Sleep(RETRY_DELAY_MS * retry); // Exponential backoff
        }
        
        // Try to open clipboard with a timeout approach
        // Use a very short timeout by trying multiple times quickly
        bool opened = false;
        for (int quickRetry = 0; quickRetry < 10; quickRetry++) {
            if (OpenClipboard(hwndMain)) {
                opened = true;
                break;
            }
            Sleep(5); // Very short delay
        }
        
        if (!opened) {
            continue; // Try again in next iteration
        }
        
        // Successfully opened clipboard - process it quickly
        try {
            // Determine primary format (for display) - process in priority order
            UINT priorityFormats[] = {
                CF_HDROP,        // Files
                CF_UNICODETEXT,  // Unicode text
                CF_TEXT,         // ANSI text
                CF_BITMAP,       // Bitmap
                CF_DIBV5,        // DIB v5
                CF_DIB,          // DIB
                CF_ENHMETAFILE,  // Enhanced metafile
                CF_METAFILEPICT  // Metafile picture
            };
            
            UINT primaryFormat = 0;
            for (UINT priorityFormat : priorityFormats) {
                if (IsClipboardFormatAvailable(priorityFormat)) {
                    primaryFormat = priorityFormat;
                    break;
                }
            }
            
            // If no priority format found, get first available format
            if (primaryFormat == 0) {
                primaryFormat = EnumClipboardFormats(0);
            }
            
            if (primaryFormat != 0) {
                std::vector<BYTE> data;
                UINT storageFormat = primaryFormat;
                
                // Check if primary format is handle-based
                bool isPrimaryHandleBased = (
                    primaryFormat == CF_BITMAP ||
                    primaryFormat == CF_PALETTE ||
                    primaryFormat == CF_METAFILEPICT ||
                    primaryFormat == CF_ENHMETAFILE ||
                    primaryFormat == 0x0082 ||  // CF_DSPBITMAP
                    primaryFormat == 0x008E ||  // CF_DSPENHMETAFILE
                    primaryFormat == 0x0083     // CF_DSPMETAFILEPICT
                );
                
                // CF_BITMAP is handle-based, not memory-based - convert to DIB
                if (primaryFormat == CF_BITMAP) {
                    HBITMAP hBitmap = (HBITMAP)GetClipboardData(CF_BITMAP);
                    if (hBitmap) {
                        data = ConvertBitmapToDIB(hBitmap);
                        if (!data.empty()) {
                            storageFormat = CF_DIB; // Store as DIB instead of CF_BITMAP
                        } else {
                            // Conversion failed, skip this item
                            CloseClipboard();
                            isProcessingClipboard = false;
                            return;
                        }
                    } else {
                        CloseClipboard();
                        isProcessingClipboard = false;
                        return;
                    }
                } else if (isPrimaryHandleBased) {
                    // Other handle-based formats - skip them as we can't easily convert
                    CloseClipboard();
                    isProcessingClipboard = false;
                    return;
                } else {
                    // Get primary format data (for memory-based formats)
                    try {
                        HGLOBAL hMem = GetClipboardData(primaryFormat);
                        if (hMem) {
                            SIZE_T size = GlobalSize(hMem);
                            if (size > 0 && size < 100 * 1024 * 1024) { // Limit to 100MB
                                void* pMem = GlobalLock(hMem);
                                if (pMem) {
                                    try {
                                        data.resize(size);
                                        memcpy(data.data(), pMem, size);
                                        GlobalUnlock(hMem);
                                    } catch (...) {
                                        // If memcpy fails, unlock and abort
                                        GlobalUnlock(hMem);
                                        CloseClipboard();
                                        isProcessingClipboard = false;
                                        return;
                                    }
                                } else {
                                    CloseClipboard();
                                    isProcessingClipboard = false;
                                    return;
                                }
                            } else {
                                CloseClipboard();
                                isProcessingClipboard = false;
                                return;
                            }
                        } else {
                            CloseClipboard();
                            isProcessingClipboard = false;
                            return;
                        }
                    } catch (...) {
                        // Error accessing clipboard data, abort
                        CloseClipboard();
                        isProcessingClipboard = false;
                        return;
                    }
                }
                
                if (!data.empty()) {
                    // Create clipboard item with storage format (which may be different from primary if converted)
                    // Use storageFormat instead of primaryFormat so the item knows it's DIB data, not CF_BITMAP
                    std::unique_ptr<ClipboardItem> item;
                    try {
                        // Validate data size before creating item
                        if (data.size() > 0 && data.size() < 200 * 1024 * 1024) { // Limit to 200MB
                            // Use storageFormat so converted CF_BITMAP->CF_DIB items are created with CF_DIB format
                            item = std::make_unique<ClipboardItem>(storageFormat, data);
                            if (!item) {
                                CloseClipboard();
                                isProcessingClipboard = false;
                                return;
                            }
                            
                            // If we converted CF_BITMAP to DIB, also store the original format for compatibility
                            if (primaryFormat == CF_BITMAP && storageFormat == CF_DIB) {
                                // The item is already created with CF_DIB format and data
                                // No need to add it again, it's already there
                            }
                        } else {
                            // Data too large or invalid, skip
                            CloseClipboard();
                            isProcessingClipboard = false;
                            return;
                        }
                    } catch (const std::bad_alloc&) {
                        // Out of memory, skip this item
                        CloseClipboard();
                        isProcessingClipboard = false;
                        return;
                    } catch (const std::exception&) {
                        // Standard exception, skip this item
                        CloseClipboard();
                        isProcessingClipboard = false;
                        return;
                    } catch (...) {
                        // Any other error creating item, abort
                        CloseClipboard();
                        isProcessingClipboard = false;
                        return;
                    }
                    
                    // Store ALL available formats to preserve formatting
                    // CRITICAL: Limit formats aggressively to prevent memory crashes
                    int formatCount = 0;
                    size_t totalFormatSize = 0;
                    const int MAX_FORMATS_PER_ITEM = 12;
                    const size_t MAX_FORMAT_SIZE_PER_ITEM = 10 * 1024 * 1024; // 10MB per item total
                    
                    UINT format = EnumClipboardFormats(0);
                    while (format != 0 && formatCount < MAX_FORMATS_PER_ITEM) {
                        // Skip CF_BITMAP (we already converted it) and already stored primary format
                        // CRITICAL: Skip handle-based formats that can't be treated as HGLOBAL
                        // These formats cause crashes when we try to GlobalSize/GlobalLock them
                        bool isHandleBasedFormat = (
                            format == CF_BITMAP ||
                            format == CF_PALETTE ||
                            format == CF_METAFILEPICT ||
                            format == CF_ENHMETAFILE ||
                            format == 0x0082 ||  // CF_DSPBITMAP
                            format == 0x008E ||  // CF_DSPENHMETAFILE  
                            format == 0x0083     // CF_DSPMETAFILEPICT
                        );
                        
                        if (format != primaryFormat && !isHandleBasedFormat) {
                            try {
                                // Use GetClipboardData to get handle
                                HANDLE hFormatData = GetClipboardData(format);
                                if (hFormatData) {
                                    // Try to get size - this will fail for handle-based formats
                                    SIZE_T formatSize = 0;
                                    try {
                                        formatSize = GlobalSize((HGLOBAL)hFormatData);
                                    } catch (...) {
                                        // GlobalSize failed, skip this format
                                        format = EnumClipboardFormats(format);
                                        continue;
                                    }
                                    
                                    // CRITICAL: Very strict limits to prevent crashes
                                    if (formatSize > 0 && formatSize < 5 * 1024 * 1024 && // Max 5MB per format
                                        totalFormatSize + formatSize < MAX_FORMAT_SIZE_PER_ITEM) {
                                        void* pFormatMem = nullptr;
                                        try {
                                            pFormatMem = GlobalLock((HGLOBAL)hFormatData);
                                        } catch (...) {
                                            // GlobalLock failed, skip this format
                                            format = EnumClipboardFormats(format);
                                            continue;
                                        }
                                        
                                        if (pFormatMem) {
                                            try {
                                                std::vector<BYTE> formatData(formatSize);
                                                memcpy(formatData.data(), pFormatMem, formatSize);
                                                GlobalUnlock((HGLOBAL)hFormatData);
                                                item->AddFormat(format, formatData);
                                                totalFormatSize += formatSize;
                                                formatCount++;
                                            } catch (...) {
                                                // If memcpy fails, unlock and skip this format
                                                try {
                                                    GlobalUnlock((HGLOBAL)hFormatData);
                                                } catch (...) {
                                                    // Ignore unlock errors
                                                }
                                                // Continue to next format instead of breaking
                                            }
                                        }
                                    }
                                }
                            } catch (...) {
                                // Skip this format if there's any error - continue to next format
                            }
                        }
                        
                        // Get next format - wrapped in try-catch to handle Excel's weird formats
                        try {
                            format = EnumClipboardFormats(format);
                        } catch (...) {
                            break; // Stop enumeration if EnumClipboardFormats crashes
                        }
                    }
                    
                    // Check if this matches the text we just pasted (multi-paste combined text)
                    bool isPastPaste = false;
                    if (!lastPastedText.empty() && item && item->format == CF_UNICODETEXT) {
                        try {
                            const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
                            if (textData && textData->size() >= sizeof(wchar_t)) {
                                size_t len = textData->size() / sizeof(wchar_t);
                                if (len > 0) {
                                    const wchar_t* textPtr = (const wchar_t*)textData->data();
                                    size_t actualLen = len;
                                    for (size_t j = 0; j < len; j++) {
                                        if (textPtr[j] == L'\0') {
                                            actualLen = j;
                                            break;
                                        }
                                    }
                                    if (actualLen > 0) {
                                        std::wstring clipboardText(textPtr, actualLen);
                                        if (clipboardText == lastPastedText) {
                                            isPastPaste = true;
                                            // Clear lastPastedText after matching to allow future pastes
                                            lastPastedText.clear();
                                        }
                                    }
                                }
                            }
                        } catch (...) {
                            // Error comparing, treat as not matching
                        }
                    }
                    
                    // Check if this is a duplicate of the last item
                    bool isDuplicate = false;
                    if (!clipboardHistory.empty() && item) {
                        try {
                            // Make a copy of the reference to avoid invalidation issues
                            const auto& lastItem = clipboardHistory[0];
                            if (lastItem && item->formats.size() > 0 && lastItem->formats.size() > 0) {
                                // Compare formats - must have same number of formats
                                if (item->formats.size() == lastItem->formats.size()) {
                                    // Check if all formats match
                                    bool allFormatsMatch = true;
                                    const size_t MAX_COMPARE_SIZE = 10 * 1024 * 1024; // Limit comparison to 10MB per format
                                    
                                    for (const auto& formatPair : item->formats) {
                                        UINT fmt = formatPair.first;
                                        const std::vector<BYTE>& newData = formatPair.second;
                                        
                                        // Skip detailed comparison for very large formats (like RTF) - compare only text formats
                                        // This prevents crashes when Word includes large RTF data
                                        // Only do detailed comparison for text formats and the primary format
                                        bool isTextFormat = (fmt == CF_UNICODETEXT || fmt == CF_TEXT || fmt == CF_OEMTEXT);
                                        bool isPrimaryFormat = (fmt == item->format || fmt == lastItem->format);
                                        
                                        if (!isTextFormat && !isPrimaryFormat) {
                                            // For non-text, non-primary formats, just check if format exists
                                            const std::vector<BYTE>* lastData = lastItem->GetFormatData(fmt);
                                            if (!lastData) {
                                                allFormatsMatch = false;
                                                break;
                                            }
                                            continue; // Skip detailed comparison for non-primary formats
                                        }
                                        
                                        // Validate newData
                                        if (newData.empty() || newData.size() > MAX_COMPARE_SIZE) {
                                            allFormatsMatch = false;
                                            break;
                                        }
                                        
                                        const std::vector<BYTE>* lastData = lastItem->GetFormatData(fmt);
                                        if (!lastData || lastData->empty() || lastData->size() != newData.size() || lastData->size() > MAX_COMPARE_SIZE) {
                                            allFormatsMatch = false;
                                            break;
                                        }
                                        
                                        // Validate pointers and size before memcmp
                                        size_t compareSize = std::min(newData.size(), lastData->size());
                                        if (compareSize > 0 && compareSize <= MAX_COMPARE_SIZE && 
                                            newData.data() && lastData->data()) {
                                            // Compare data byte-by-byte
                                            if (memcmp(newData.data(), lastData->data(), compareSize) != 0) {
                                                allFormatsMatch = false;
                                                break;
                                            }
                                        } else {
                                            allFormatsMatch = false;
                                            break;
                                        }
                                    }
                                    
                                    if (allFormatsMatch) {
                                        isDuplicate = true;
                                    }
                                }
                            }
                            // Same text copied twice (e.g. Ctrl+C twice on "text") → treat as duplicate even if raw bytes differ
                            if (!isDuplicate && item) {
                                std::wstring newNorm = GetNormalizedTextForDuplicateCheck(item.get());
                                std::wstring lastNorm = GetNormalizedTextForDuplicateCheck(clipboardHistory[0].get());
                                if (!newNorm.empty() && newNorm == lastNorm)
                                    isDuplicate = true;
                            }
                        } catch (...) {
                            // Error during duplicate check, treat as not duplicate and continue
                            isDuplicate = false;
                        }
                    }
                    
                    // Only add if not a duplicate and not something we just pasted
                    if (!isDuplicate && !isPastPaste && item) {
                        try {
                            // Add to front of list (most recent first)
                            clipboardHistory.insert(
                                clipboardHistory.begin(),
                                std::move(item)
                            );
                            
                            // Limit to maxItems (never evicting pinned items)
                            TrimHistory();
                            
                            // Update filter if list is visible
                            if (listVisible) {
                                FilterItems();
                                UpdateListWindow();
                            }
                        } catch (...) {
                            // Error adding item to history, skip it
                            // Don't crash, just continue
                        }
                    }
                }
            }
            
            // Close clipboard immediately after processing
            CloseClipboard();
            
            // Successfully processed, exit retry loop
            isProcessingClipboard = false;
            return;
        } catch (...) {
            // Error occurred, ensure clipboard is closed and reset flag
            // Note: CloseClipboard can be called even if clipboard wasn't opened
            try {
                CloseClipboard();
            } catch (...) {
                // Ignore errors closing clipboard
            }
            isProcessingClipboard = false;
        }
    }
    
    // Finished all retries, ensure clipboard is closed and reset flag
    // Note: CloseClipboard can be called even if clipboard wasn't opened
    try {
        CloseClipboard();
    } catch (...) {
        // Ignore errors closing clipboard
    }
    isProcessingClipboard = false;
}

// Resource ID for embedded click.mp3 (must match ClipboardManager.rc)
#define IDR_CLICK_MP3 101

// Get path to click.mp3: first try embedded resource (extract to temp), then exe directory.
static std::wstring GetClickMp3PathFromResource() {
    static std::wstring s_path;
    static bool s_tried = false;
    if (s_tried) return s_path;
    s_tried = true;

    HMODULE hMod = GetModuleHandleW(nullptr);
    HRSRC hRes = FindResourceW(hMod, MAKEINTRESOURCEW(IDR_CLICK_MP3), RT_RCDATA);
    if (hRes) {
        HGLOBAL hGlob = LoadResource(hMod, hRes);
        if (hGlob) {
            const void* pData = LockResource(hGlob);
            if (pData) {
                DWORD size = SizeofResource(hMod, hRes);
                if (size > 0) {
                    wchar_t tempDir[MAX_PATH];
                    if (GetTempPathW(MAX_PATH, tempDir) != 0) {
                        s_path = std::wstring(tempDir) + L"clip2_click.mp3";
                        HANDLE hFile = CreateFileW(s_path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                        if (hFile != INVALID_HANDLE_VALUE) {
                            DWORD written = 0;
                            BOOL ok = WriteFile(hFile, pData, size, &written, nullptr);
                            CloseHandle(hFile);
                            if (ok && written == size)
                                return s_path;
                        }
                        s_path.clear();
                    }
                }
            }
        }
    }

    // Fallback: click.mp3 next to the executable
    wchar_t exePath[MAX_PATH];
    if (GetModuleFileNameW(hMod, exePath, MAX_PATH) != 0) {
        std::wstring dir = exePath;
        size_t lastSlash = dir.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos)
            dir.resize(lastSlash + 1);
        dir += L"click.mp3";
        DWORD att = GetFileAttributesW(dir.c_str());
        if (att != INVALID_FILE_ATTRIBUTES && !(att & FILE_ATTRIBUTE_DIRECTORY))
            s_path = dir;
    }
    return s_path;
}

// Decode click.mp3 to a temporary WAV using Media Foundation; return path to WAV or empty on failure.
// PlaySound(WAV file) is very reliable, so we prefer this over MCI/WMP for the copy sound.
static std::wstring GetOrCreateClickWavPath() {
    static std::wstring s_wavPath;
    static bool s_succeeded = false;
    if (s_succeeded) return s_wavPath;

    std::wstring mp3Path = GetClickMp3PathFromResource();
    if (mp3Path.empty()) return s_wavPath;

    HRESULT hr = MFStartup(MF_VERSION, MFSTARTUP_NOSOCKET);
    if (FAILED(hr)) return s_wavPath;

    IMFSourceReader* pReader = nullptr;
    std::wstring url = L"file:///";
    for (wchar_t c : mp3Path) {
        if (c == L'\\') url += L'/';
        else if (c == L' ') url += L"%20";
        else url += c;
    }
    hr = MFCreateSourceReaderFromURL(url.c_str(), nullptr, &pReader);
    if (FAILED(hr) || !pReader) { MFShutdown(); return s_wavPath; }

    IMFMediaType* pPCMType = nullptr;
    hr = MFCreateMediaType(&pPCMType);
    if (FAILED(hr)) { pReader->Release(); MFShutdown(); return s_wavPath; }
    pPCMType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio);
    pPCMType->SetGUID(MF_MT_SUBTYPE, MFAudioFormat_PCM);
    hr = pReader->SetCurrentMediaType((DWORD)MF_SOURCE_READER_FIRST_AUDIO_STREAM, nullptr, pPCMType);
    pPCMType->Release();
    if (FAILED(hr)) { pReader->Release(); MFShutdown(); return s_wavPath; }

    IMFMediaType* pOutType = nullptr;
    hr = pReader->GetCurrentMediaType((DWORD)MF_SOURCE_READER_FIRST_AUDIO_STREAM, &pOutType);
    if (FAILED(hr) || !pOutType) { pReader->Release(); MFShutdown(); return s_wavPath; }
    UINT32 sampleRate = 44100, channels = 1;
    pOutType->GetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND, &sampleRate);
    pOutType->GetUINT32(MF_MT_AUDIO_NUM_CHANNELS, &channels);
    pOutType->Release();

    std::vector<BYTE> pcmData;
    IMFSample* pSample = nullptr;
    IMFMediaBuffer* pBuffer = nullptr;
    while (true) {
        DWORD flags = 0;
        LONGLONG ts = 0;
        hr = pReader->ReadSample((DWORD)MF_SOURCE_READER_FIRST_AUDIO_STREAM, 0, nullptr, &flags, &ts, &pSample);
        if (FAILED(hr) || (flags & MF_SOURCE_READERF_ENDOFSTREAM)) break;
        if (!pSample) continue;
        hr = pSample->ConvertToContiguousBuffer(&pBuffer);
        pSample->Release();
        if (FAILED(hr) || !pBuffer) continue;
        BYTE* pData = nullptr;
        DWORD maxLen = 0, dataLen = 0;
        hr = pBuffer->Lock(&pData, &maxLen, &dataLen);
        if (SUCCEEDED(hr) && pData && dataLen > 0) {
            pcmData.insert(pcmData.end(), pData, pData + dataLen);
            pBuffer->Unlock();
        }
        pBuffer->Release();
    }
    pReader->Release();
    MFShutdown();
    if (pcmData.empty()) return s_wavPath;

    wchar_t tempDir[MAX_PATH];
    if (GetTempPathW(MAX_PATH, tempDir) == 0) return s_wavPath;
    s_wavPath = std::wstring(tempDir) + L"clip2_click.wav";
    HANDLE hFile = CreateFileW(s_wavPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) return s_wavPath;

#pragma pack(push, 1)
    struct WavHeader {
        char riff[4];
        DWORD fileLen;
        char wave[4];
        char fmt[4];
        DWORD fmtLen;
        WORD format;
        WORD channels;
        DWORD sampleRate;
        DWORD byteRate;
        WORD blockAlign;
        WORD bitsPerSample;
        char data[4];
        DWORD dataLen;
    };
#pragma pack(pop)
    WavHeader h = {};
    memcpy(h.riff, "RIFF", 4);
    h.fileLen = (DWORD)(sizeof(WavHeader) - 8 + pcmData.size());
    memcpy(h.wave, "WAVE", 4);
    memcpy(h.fmt, "fmt ", 4);
    h.fmtLen = 16;
    h.format = 1;
    h.channels = (WORD)channels;
    h.sampleRate = sampleRate;
    h.bitsPerSample = 16;
    h.blockAlign = (WORD)(channels * 2);
    h.byteRate = sampleRate * h.blockAlign;
    memcpy(h.data, "data", 4);
    h.dataLen = (DWORD)pcmData.size();
    DWORD written = 0;
    WriteFile(hFile, &h, sizeof(h), &written, nullptr);
    WriteFile(hFile, pcmData.data(), (DWORD)pcmData.size(), &written, nullptr);
    CloseHandle(hFile);
    s_succeeded = true;  // Only cache success so we retry on next call if startup failed
    return s_wavPath;
}

// Play MP3 via Windows Media Player COM (used when MCI fails).
static bool PlayMp3ViaWMP(const std::wstring& path) {
    static const CLSID CLSID_WMP = {
        0x6BF52A52, 0x394A, 0x11d3, { 0xB1, 0x53, 0x00, 0xC0, 0x4F, 0x79, 0xFA, 0xA6 }
    };
    IDispatch* pWmp = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_WMP, nullptr, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pWmp);
    if (FAILED(hr) || !pWmp) return false;

    OLECHAR* urlName = const_cast<OLECHAR*>(L"URL");
    DISPID dispidUrl = 0;
    hr = pWmp->GetIDsOfNames(IID_NULL, &urlName, 1, LOCALE_USER_DEFAULT, &dispidUrl);
    if (FAILED(hr)) { pWmp->Release(); return false; }

    BSTR bstrPath = SysAllocString(path.c_str());
    if (!bstrPath) { pWmp->Release(); return false; }
    VARIANT varPath;
    VariantInit(&varPath);
    varPath.vt = VT_BSTR;
    varPath.bstrVal = bstrPath;

    DISPPARAMS dp = {};
    dp.cArgs = 1;
    dp.rgvarg = &varPath;
    dp.cNamedArgs = 1;
    DISPID dispidPut = DISPID_PROPERTYPUT;
    dp.rgdispidNamedArgs = &dispidPut;

    hr = pWmp->Invoke(dispidUrl, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, nullptr, nullptr, nullptr);
    SysFreeString(bstrPath);
    if (FAILED(hr)) { pWmp->Release(); return false; }

    OLECHAR* ctrlName = const_cast<OLECHAR*>(L"Controls");
    DISPID dispidCtrl = 0;
    hr = pWmp->GetIDsOfNames(IID_NULL, &ctrlName, 1, LOCALE_USER_DEFAULT, &dispidCtrl);
    if (FAILED(hr)) { pWmp->Release(); return false; }

    VARIANT varCtrl;
    VariantInit(&varCtrl);
    DISPPARAMS dpGet = {};
    hr = pWmp->Invoke(dispidCtrl, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpGet, &varCtrl, nullptr, nullptr);
    if (FAILED(hr) || varCtrl.vt != VT_DISPATCH || !varCtrl.pdispVal) {
        pWmp->Release();
        return false;
    }

    IDispatch* pCtrl = varCtrl.pdispVal;
    OLECHAR* playName = const_cast<OLECHAR*>(L"play");
    DISPID dispidPlay = 0;
    hr = pCtrl->GetIDsOfNames(IID_NULL, &playName, 1, LOCALE_USER_DEFAULT, &dispidPlay);
    if (FAILED(hr)) { pCtrl->Release(); pWmp->Release(); return false; }

    DISPPARAMS dpPlay = {};
    hr = pCtrl->Invoke(dispidPlay, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpPlay, nullptr, nullptr, nullptr);
    pCtrl->Release();
    pWmp->Release();
    return SUCCEEDED(hr);
}

// Fallback: short WAV generated in-memory when click.mp3 cannot be played.
static const void* GetEmbeddedClickWav(size_t* outSize) {
#pragma pack(push, 1)
    struct WavHeader {
        char riff[4];
        DWORD fileLen;
        char wave[4];
        char fmt[4];
        DWORD fmtLen;
        WORD format;
        WORD channels;
        DWORD sampleRate;
        DWORD byteRate;
        WORD blockAlign;
        WORD bitsPerSample;
        char data[4];
        DWORD dataLen;
    };
#pragma pack(pop)
    static std::vector<char> s_wav;
    static bool s_built = false;
    if (!s_built) {
        s_built = true;
        const DWORD sampleRate = 22050;
        const DWORD durationMs = 60;
        const DWORD numSamples = sampleRate * durationMs / 1000;
        const DWORD dataLen = numSamples * 2;
        const DWORD fileLen = sizeof(WavHeader) - 8 + dataLen;
        s_wav.resize(sizeof(WavHeader) + dataLen);
        WavHeader* h = (WavHeader*)s_wav.data();
        memcpy(h->riff, "RIFF", 4);
        h->fileLen = fileLen;
        memcpy(h->wave, "WAVE", 4);
        memcpy(h->fmt, "fmt ", 4);
        h->fmtLen = 16;
        h->format = 1;
        h->channels = 1;
        h->sampleRate = sampleRate;
        h->byteRate = sampleRate * 2;
        h->blockAlign = 2;
        h->bitsPerSample = 16;
        memcpy(h->data, "data", 4);
        h->dataLen = dataLen;
        short* samples = (short*)(s_wav.data() + sizeof(WavHeader));
        for (DWORD i = 0; i < numSamples; i++) {
            double t = (double)i / sampleRate;
            double env = (t < 0.01) ? 1.0 : (t < 0.06 ? 1.0 - (t - 0.01) / 0.05 : 0.0);
            double v = env * 0.35 * 32767.0 * std::sin(2.0 * 3.14159265358979 * 800.0 * t);
            samples[i] = (short)(v < -32768 ? -32768 : (v > 32767 ? 32767 : (int)v));
        }
    }
    if (outSize) *outSize = s_wav.size();
    return s_wav.data();
}

void ClipboardManager::WarmUpClickSound() {
    GetClickMp3PathFromResource();
    std::wstring wavPath = GetOrCreateClickWavPath();
    // Play once synchronously at startup so the first copy also sounds (primes the audio pipeline)
    if (!wavPath.empty())
        PlaySoundW(wavPath.c_str(), nullptr, SND_FILENAME | SND_SYNC | SND_NODEFAULT);
}

void ClipboardManager::PlayClickSound() {
    // 1) Media Foundation: decode click.mp3 to WAV once, then PlaySound(WAV) — most reliable
    std::wstring wavPath = GetOrCreateClickWavPath();
    if (!wavPath.empty()) {
        if (PlaySoundW(wavPath.c_str(), nullptr, SND_FILENAME | SND_ASYNC | SND_NODEFAULT))
            return;
    }
    // 2) Try click.wav next to exe (no decode needed)
    wchar_t exePath[MAX_PATH];
    if (GetModuleFileNameW(nullptr, exePath, MAX_PATH) != 0) {
        std::wstring dir = exePath;
        size_t lastSlash = dir.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) dir.resize(lastSlash + 1);
        std::wstring localWav = dir + L"click.wav";
        if (GetFileAttributesW(localWav.c_str()) != INVALID_FILE_ATTRIBUTES) {
            if (PlaySoundW(localWav.c_str(), nullptr, SND_FILENAME | SND_ASYNC | SND_NODEFAULT))
                return;
        }
    }
    // 3) MCI with click.mp3
    std::wstring mp3Path = GetClickMp3PathFromResource();
    if (!mp3Path.empty()) {
        static bool aliasOpen = false;
        if (aliasOpen) {
            mciSendString(L"stop clicksnd", nullptr, 0, nullptr);
            mciSendString(L"close clicksnd", nullptr, 0, nullptr);
            aliasOpen = false;
        }
        std::wstring openCmd = L"open \"" + mp3Path + L"\" type MPEGVideo alias clicksnd";
        MCIERROR err = mciSendString(openCmd.c_str(), nullptr, 0, nullptr);
        if (err != 0) {
            openCmd = L"open \"" + mp3Path + L"\" alias clicksnd";
            err = mciSendString(openCmd.c_str(), nullptr, 0, nullptr);
        }
        if (err == 0) {
            aliasOpen = true;
            mciSendString(L"play clicksnd", nullptr, 0, nullptr);
            return;
        }
        // 4) WMP for click.mp3
        try {
            if (PlayMp3ViaWMP(mp3Path))
                return;
        } catch (...) {}
    }
    // 5) Fallback: embedded WAV so something always plays
    size_t len = 0;
    const void* wav = GetEmbeddedClickWav(&len);
    if (wav && len > 0)
        PlaySoundA((LPCSTR)wav, nullptr, SND_MEMORY | SND_ASYNC | SND_NODEFAULT);
}

// Enforce the maxItems cap by removing the oldest UNPINNED items first, so pinned
// favorites are never silently evicted even when history is full.
void ClipboardManager::TrimHistory() {
    int cap = maxItems;
    if (cap < MIN_MAX_ITEMS) cap = MIN_MAX_ITEMS;
    if (cap > MAX_MAX_ITEMS) cap = MAX_MAX_ITEMS;
    while ((int)clipboardHistory.size() > cap) {
        // Find the last (oldest) unpinned item to remove.
        int removeIdx = -1;
        for (int i = (int)clipboardHistory.size() - 1; i >= 0; i--) {
            if (clipboardHistory[i] && !clipboardHistory[i]->pinned) { removeIdx = i; break; }
        }
        if (removeIdx < 0) break;  // everything is pinned; nothing to trim
        clipboardHistory.erase(clipboardHistory.begin() + removeIdx);
    }
}

void ClipboardManager::TogglePin(int filteredIndex) {
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) return;
    int actualIndex = filteredIndices[filteredIndex];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) return;
    if (!clipboardHistory[actualIndex]) return;
    clipboardHistory[actualIndex]->pinned = !clipboardHistory[actualIndex]->pinned;
    SaveClipboardHistory();
    FilterItems();
    UpdateListWindow();
}

void ClipboardManager::CopyItemAsPlainText(int filteredIndex) {
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) return;
    int actualIndex = filteredIndices[filteredIndex];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) return;
    std::wstring text = GetPlainTextForDirectPaste(clipboardHistory[actualIndex].get());
    if (text.empty()) return;
    lastPastedText = text;  // avoid re-capturing our own clipboard write
    SetClipboardUnicodeOnly(hwndMain, text);
    PlayClickSound();
}

void ClipboardManager::OpenItemUrl(int filteredIndex) {
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) return;
    int actualIndex = filteredIndices[filteredIndex];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) return;
    std::wstring text = GetPlainTextForDirectPaste(clipboardHistory[actualIndex].get());
    // Trim surrounding whitespace.
    size_t a = text.find_first_not_of(L" \t\r\n");
    size_t b = text.find_last_not_of(L" \t\r\n");
    if (a == std::wstring::npos) return;
    std::wstring url = text.substr(a, b - a + 1);
    if (url.rfind(L"http://", 0) != 0 && url.rfind(L"https://", 0) != 0) return;
    ShellExecuteW(nullptr, L"open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

void ClipboardManager::SaveImageItemToFile(int filteredIndex) {
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) return;
    int actualIndex = filteredIndices[filteredIndex];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) return;
    const auto& item = clipboardHistory[actualIndex];
    if (!item || !item->isImage) return;

    // Prefer DIB bytes (which are the BITMAPINFO + pixel data we can wrap into a .bmp file).
    const std::vector<BYTE>* dib = item->GetFormatData(CF_DIB);
    if (!dib || dib->empty()) dib = item->GetFormatData(CF_DIBV5);
    if (!dib || dib->size() < sizeof(BITMAPINFOHEADER)) {
        MessageBoxW(hwndList, L"This image cannot be exported (no DIB data).", L"clip2", MB_OK | MB_ICONINFORMATION);
        return;
    }

    wchar_t fileName[MAX_PATH] = L"clip2_image.bmp";
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwndList;
    ofn.lpstrFilter = L"Bitmap Image (*.bmp)\0*.bmp\0All Files\0*.*\0";
    ofn.lpstrFile = fileName;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrDefExt = L"bmp";
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
    if (!GetSaveFileNameW(&ofn)) return;

    // Compose a BMP file: BITMAPFILEHEADER + DIB bytes. Compute pixel-array offset
    // from the header (accounting for color table / bitfields where present).
    const BITMAPINFOHEADER* bih = (const BITMAPINFOHEADER*)dib->data();
    DWORD headerSize = bih->biSize ? bih->biSize : sizeof(BITMAPINFOHEADER);
    DWORD clrUsed = bih->biClrUsed;
    if (clrUsed == 0 && bih->biBitCount <= 8) clrUsed = (DWORD)1 << bih->biBitCount;
    DWORD paletteBytes = clrUsed * sizeof(RGBQUAD);
    if (bih->biCompression == BI_BITFIELDS) paletteBytes += 3 * sizeof(DWORD);
    DWORD pixelOffset = sizeof(BITMAPFILEHEADER) + headerSize + paletteBytes;

    BITMAPFILEHEADER bfh = {};
    bfh.bfType = 0x4D42;  // 'BM'
    bfh.bfOffBits = pixelOffset;
    bfh.bfSize = sizeof(BITMAPFILEHEADER) + (DWORD)dib->size();

    HANDLE h = CreateFileW(fileName, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) {
        MessageBoxW(hwndList, L"Could not create the file.", L"clip2", MB_OK | MB_ICONERROR);
        return;
    }
    DWORD written = 0;
    WriteFile(h, &bfh, sizeof(bfh), &written, nullptr);
    WriteFile(h, dib->data(), (DWORD)dib->size(), &written, nullptr);
    CloseHandle(h);
}

void ClipboardManager::DeleteItem(int filteredIndex) {
    // Validate filtered index
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) {
        return;
    }
    
    // Get the actual index in clipboardHistory
    int actualIndex = filteredIndices[filteredIndex];
    
    // Validate actual index
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) {
        return;
    }
    
    // Remove the item from clipboardHistory
    clipboardHistory.erase(clipboardHistory.begin() + actualIndex);
    
    // Rebuild filteredIndices - adjust all indices that were after the deleted item
    // and filter again to account for the removed item
    FilterItems();
    
    // Update UI state
    if (filteredIndices.empty()) {
        selectedIndex = -1;
        scrollOffset = 0;
    } else {
        // Adjust selectedIndex if necessary
        if (selectedIndex >= (int)filteredIndices.size()) {
            selectedIndex = (int)filteredIndices.size() - 1;
        }
        if (selectedIndex < 0 && !filteredIndices.empty()) {
            selectedIndex = 0;
        }
        
        // Adjust scroll offset if necessary
        int maxScroll = std::max(0, (int)filteredIndices.size() - itemsPerPage);
        if (scrollOffset > maxScroll) {
            scrollOffset = maxScroll;
        }
    }
    
    hoveredItemIndex = -1;
    HidePreviewWindow();
    
    // Update the display
    UpdateListWindow();
}

bool ClipboardManager::IsTextItem(int actualIndex) {
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) {
        return false;
    }
    
    const auto& item = clipboardHistory[actualIndex];
    // Check if item has text formats
    return (item->format == CF_TEXT || item->format == CF_UNICODETEXT || item->format == CF_OEMTEXT ||
            item->GetFormatData(CF_TEXT) != nullptr || 
            item->GetFormatData(CF_UNICODETEXT) != nullptr ||
            item->GetFormatData(CF_OEMTEXT) != nullptr);
}

void ClipboardManager::TransformTextItem(int filteredIndex, int transformType) {
    // Validate filtered index
    if (filteredIndex < 0 || filteredIndex >= (int)filteredIndices.size()) {
        return;
    }
    
    int actualIndex = filteredIndices[filteredIndex];
    if (!IsTextItem(actualIndex)) {
        return; // Not a text item
    }
    
    const auto& item = clipboardHistory[actualIndex];
    
    // Extract text - prefer Unicode text
    std::wstring text;
    const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
    if (textData && textData->size() >= sizeof(wchar_t)) {
        size_t len = textData->size() / sizeof(wchar_t);
        if (len > 0) {
            const wchar_t* textPtr = (const wchar_t*)textData->data();
            // Handle null-terminated string - find null terminator or use full length
            size_t actualLen = len;
            for (size_t i = 0; i < len; i++) {
                if (textPtr[i] == L'\0') {
                    actualLen = i;
                    break;
                }
            }
            if (actualLen > 0) {
                text = std::wstring(textPtr, actualLen);
            }
        }
    } else {
        // Try ANSI text
        textData = item->GetFormatData(CF_TEXT);
        if (textData && textData->size() > 0) {
            size_t len = textData->size();
            const char* ansiText = (const char*)textData->data();
            // Find null terminator
            size_t actualLen = len;
            for (size_t i = 0; i < len; i++) {
                if (ansiText[i] == '\0') {
                    actualLen = i;
                    break;
                }
            }
            if (actualLen > 0) {
                std::string ansiStr(ansiText, actualLen);
                int wideLen = MultiByteToWideChar(CP_ACP, 0, ansiStr.c_str(), (int)ansiStr.length(), nullptr, 0);
                if (wideLen > 0) {
                    text.resize(wideLen);
                    MultiByteToWideChar(CP_ACP, 0, ansiStr.c_str(), (int)ansiStr.length(), &text[0], wideLen);
                }
            }
        } else {
            // Try OEM text
            textData = item->GetFormatData(CF_OEMTEXT);
            if (textData && textData->size() > 0) {
                size_t len = textData->size();
                const char* oemText = (const char*)textData->data();
                // Find null terminator
                size_t actualLen = len;
                for (size_t i = 0; i < len; i++) {
                    if (oemText[i] == '\0') {
                        actualLen = i;
                        break;
                    }
                }
                if (actualLen > 0) {
                    std::string oemStr(oemText, actualLen);
                    int wideLen = MultiByteToWideChar(CP_OEMCP, 0, oemStr.c_str(), (int)oemStr.length(), nullptr, 0);
                    if (wideLen > 0) {
                        text.resize(wideLen);
                        MultiByteToWideChar(CP_OEMCP, 0, oemStr.c_str(), (int)oemStr.length(), &text[0], wideLen);
                    }
                }
            }
        }
    }
    
    if (text.empty()) {
        return; // No text found
    }
    
    // Apply transformation
    std::wstring transformedText = text;
    
    switch (transformType) {
        case TRANSFORM_UPPERCASE: {
            std::transform(transformedText.begin(), transformedText.end(), transformedText.begin(), ::towupper);
            break;
        }
        case TRANSFORM_LOWERCASE: {
            std::transform(transformedText.begin(), transformedText.end(), transformedText.begin(), ::towlower);
            break;
        }
        case TRANSFORM_TITLE_CASE: {
            bool newWord = true;
            for (size_t i = 0; i < transformedText.length(); i++) {
                if (iswspace(transformedText[i]) || iswpunct(transformedText[i])) {
                    newWord = true;
                } else if (newWord) {
                    transformedText[i] = towupper(transformedText[i]);
                    newWord = false;
                } else {
                    transformedText[i] = towlower(transformedText[i]);
                }
            }
            break;
        }
        case TRANSFORM_REMOVE_LINE_BREAKS: {
            std::wstring result;
            result.reserve(transformedText.length());
            for (wchar_t c : transformedText) {
                if (c != L'\r' && c != L'\n') {
                    result += c;
                } else {
                    // Replace with space if previous character isn't already a space
                    if (!result.empty() && result.back() != L' ') {
                        result += L' ';
                    }
                }
            }
            transformedText = result;
            break;
        }
        case TRANSFORM_TRIM_WHITESPACE: {
            // Trim leading whitespace
            size_t start = transformedText.find_first_not_of(L" \t\r\n");
            if (start != std::wstring::npos) {
                transformedText = transformedText.substr(start);
            } else {
                transformedText.clear();
            }
            // Trim trailing whitespace
            size_t end = transformedText.find_last_not_of(L" \t\r\n");
            if (end != std::wstring::npos) {
                transformedText = transformedText.substr(0, end + 1);
            } else {
                transformedText.clear();
            }
            break;
        }
        case TRANSFORM_PLAIN_TEXT: {
            // Same as TRANSFORM_PLAIN_TEXT - just use the text as-is (already extracted)
            // This transformation removes formatting by only keeping plain text
            break;
        }
        default:
            return; // Unknown transformation type
    }
    
    // Update the item with transformed text
    // Remove all text formats and add only Unicode text
    clipboardHistory[actualIndex]->formats.erase(CF_TEXT);
    clipboardHistory[actualIndex]->formats.erase(CF_OEMTEXT);
    clipboardHistory[actualIndex]->formats.erase(CF_UNICODETEXT);
    
    // Add transformed text as Unicode
    size_t dataSize = (transformedText.length() + 1) * sizeof(wchar_t);
    std::vector<BYTE> unicodeData(dataSize);
    wchar_t* textPtr = (wchar_t*)unicodeData.data();
    wcscpy_s(textPtr, transformedText.length() + 1, transformedText.c_str());
    
    clipboardHistory[actualIndex]->formats[CF_UNICODETEXT] = unicodeData;
    clipboardHistory[actualIndex]->format = CF_UNICODETEXT;
    
    // Update preview
    clipboardHistory[actualIndex]->preview = transformedText.length() > 50 
        ? transformedText.substr(0, 50) + L"..." 
        : transformedText;
    
    // Copy transformed text to clipboard
    // Store the transformed text so we can ignore it if it's copied back
    lastPastedText = transformedText;
    
    isPasting = true;
    Sleep(10);
    
    if (OpenClipboard(hwndMain)) {
        EmptyClipboard();
        
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize);
        if (hMem) {
            wchar_t* pMem = (wchar_t*)GlobalLock(hMem);
            if (pMem) {
                wcscpy_s(pMem, transformedText.length() + 1, transformedText.c_str());
                GlobalUnlock(hMem);
                
                if (SetClipboardData(CF_UNICODETEXT, hMem)) {
                    lastSequenceNumber = GetClipboardSequenceNumber();
                } else {
                    GlobalFree(hMem);
                }
            } else {
                GlobalFree(hMem);
            }
        }
        
        CloseClipboard();
    }
    
    // Keep isPasting flag true longer to prevent re-adding the transformed item
    Sleep(500); // Delay to ignore clipboard updates from transform operation
    
    isPasting = false;
    
    // Update the display to show the transformed preview
    FilterItems(); // Rebuild filtered indices in case search is active
    UpdateListWindow(); // Refresh the UI to show updated preview
}

void ClipboardManager::ClearClipboardHistory() {
    // Preserve pinned/favorite items; everything else is removed.
    std::vector<std::unique_ptr<ClipboardItem>> kept;
    for (auto& item : clipboardHistory) {
        if (item && item->pinned) kept.push_back(std::move(item));
    }
    clipboardHistory.swap(kept);

    // Reset related UI state
    scrollOffset = 0;
    hoveredItemIndex = -1;
    selectedIndex = clipboardHistory.empty() ? -1 : 0;
    numberInput.clear();
    searchText.clear();
    if (hwndSearch) {
        SetWindowText(hwndSearch, L"");
    }
    FilterItems();
    UpdateListWindow();
}

// ----- history.dat persistence helpers -----------------------------------
//
// On-disk container:
//   "CLP3" + DWORD blobLen + DPAPI-encrypted blob   (current; encrypted at rest)
// The decrypted blob (also the legacy plaintext file) is an inner payload:
//   "CLP2" + DWORD version + DWORD count + records
//     version 1 record: DWORD fmt, DWORD size, bytes   (text only; legacy)
//     version 2 record: BYTE pinned, DWORD formatCount, [DWORD fmt, DWORD size, bytes]*
namespace {
    const DWORD kPersistFormatCap = 16 * 1024 * 1024;  // per-format byte cap
    const DWORD kPersistItemCap   = 16 * 1024 * 1024;  // per-item total byte cap

    void AppendBytes(std::vector<BYTE>& buf, const void* p, size_t n) {
        const BYTE* b = (const BYTE*)p;
        buf.insert(buf.end(), b, b + n);
    }
    void AppendDword(std::vector<BYTE>& buf, DWORD v) { AppendBytes(buf, &v, sizeof(DWORD)); }

    bool DpapiProtect(const std::vector<BYTE>& in, std::vector<BYTE>& out) {
        DATA_BLOB din; din.cbData = (DWORD)in.size(); din.pbData = const_cast<BYTE*>(in.data());
        DATA_BLOB dout; dout.cbData = 0; dout.pbData = nullptr;
        if (!CryptProtectData(&din, L"clip2 history", nullptr, nullptr, nullptr, 0, &dout))
            return false;
        out.assign(dout.pbData, dout.pbData + dout.cbData);
        if (dout.pbData) LocalFree(dout.pbData);
        return true;
    }
    bool DpapiUnprotect(const BYTE* data, size_t len, std::vector<BYTE>& out) {
        DATA_BLOB din; din.cbData = (DWORD)len; din.pbData = const_cast<BYTE*>(data);
        DATA_BLOB dout; dout.cbData = 0; dout.pbData = nullptr;
        if (!CryptUnprotectData(&din, nullptr, nullptr, nullptr, nullptr, 0, &dout))
            return false;
        out.assign(dout.pbData, dout.pbData + dout.cbData);
        if (dout.pbData) LocalFree(dout.pbData);
        return true;
    }

    // Build a history item from a full set of formats. Chooses a sensible primary format
    // so the existing constructor regenerates preview/thumbnail, then attaches the rest.
    std::unique_ptr<ClipboardItem> MakeItemFromFormats(std::map<UINT, std::vector<BYTE>>& fmts, bool pinned) {
        if (fmts.empty()) return nullptr;
        UINT primary = 0;
        const UINT order[] = { CF_DIBV5, CF_DIB, CF_BITMAP, CF_HDROP, CF_UNICODETEXT, CF_TEXT, CF_OEMTEXT };
        for (UINT cand : order) { if (fmts.count(cand)) { primary = cand; break; } }
        if (primary == 0) primary = fmts.begin()->first;
        std::unique_ptr<ClipboardItem> item;
        try {
            item = std::make_unique<ClipboardItem>(primary, fmts[primary]);
        } catch (...) { return nullptr; }
        if (!item || item->formats.empty()) return nullptr;  // constructor rejected the data
        for (auto& kv : fmts) {
            if (kv.first != primary && !kv.second.empty())
                item->AddFormat(kv.first, kv.second);
        }
        item->pinned = pinned;
        return item;
    }
}

void ClipboardManager::LoadClipboardHistory() {
    try {
        wchar_t appData[MAX_PATH];
        if (FAILED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, 0, appData))) return;
        std::wstring path = std::wstring(appData) + L"\\clip2";
        CreateDirectoryW(path.c_str(), nullptr);
        path += L"\\history.dat";
        HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;

        // Read the entire file (cap to a sane upper bound).
        LARGE_INTEGER fsize;
        if (!GetFileSizeEx(h, &fsize) || fsize.QuadPart <= 0 || fsize.QuadPart > (LONGLONG)256 * 1024 * 1024) {
            CloseHandle(h);
            return;
        }
        std::vector<BYTE> fileBuf((size_t)fsize.QuadPart);
        DWORD read = 0;
        if (!ReadFile(h, fileBuf.data(), (DWORD)fileBuf.size(), &read, nullptr) || read != fileBuf.size()) {
            CloseHandle(h);
            return;
        }
        CloseHandle(h);
        if (fileBuf.size() < 4) return;

        // Resolve to the inner plaintext payload (decrypt CLP3, or use CLP2 as-is for migration).
        std::vector<BYTE> payload;
        if (memcmp(fileBuf.data(), "CLP3", 4) == 0) {
            if (fileBuf.size() < 8) return;
            DWORD blobLen = 0;
            memcpy(&blobLen, fileBuf.data() + 4, 4);
            if ((size_t)blobLen + 8 > fileBuf.size()) return;
            if (!DpapiUnprotect(fileBuf.data() + 8, blobLen, payload)) return;  // can't decrypt (different user)
        } else if (memcmp(fileBuf.data(), "CLP2", 4) == 0) {
            payload = std::move(fileBuf);  // legacy plaintext, migrate on next save
        } else {
            return;
        }

        // Parse the inner payload with bounds checks.
        const BYTE* p = payload.data();
        size_t len = payload.size();
        size_t off = 0;
        auto readDword = [&](DWORD& out) -> bool {
            if (off + 4 > len) return false;
            memcpy(&out, p + off, 4); off += 4; return true;
        };
        if (len < 8 || memcmp(p, "CLP2", 4) != 0) return;
        off = 4;
        DWORD version = 0, count = 0;
        if (!readDword(version)) return;
        if (!readDword(count)) return;
        count = std::min(count, (DWORD)MAX_MAX_ITEMS);

        for (DWORD i = 0; i < count; i++) {
            if (version == 1) {
                DWORD fmt = 0, size = 0;
                if (!readDword(fmt) || !readDword(size)) break;
                if (off + size > len) break;
                if (size == 0 || size > kPersistFormatCap || (fmt != CF_UNICODETEXT && fmt != CF_TEXT)) { off += size; continue; }
                std::vector<BYTE> data(p + off, p + off + size);
                off += size;
                if (fmt == CF_UNICODETEXT) { data.push_back(0); data.push_back(0); }
                else if (fmt == CF_TEXT) { data.push_back(0); }
                std::map<UINT, std::vector<BYTE>> fmts;
                fmts[fmt] = std::move(data);
                auto item = MakeItemFromFormats(fmts, false);
                if (item) clipboardHistory.push_back(std::move(item));
            } else {  // version >= 2
                if (off + 1 > len) break;
                BYTE pinned = p[off]; off += 1;
                DWORD formatCount = 0;
                if (!readDword(formatCount)) break;
                if (formatCount > 64) break;  // corruption guard
                std::map<UINT, std::vector<BYTE>> fmts;
                bool ok = true;
                for (DWORD f = 0; f < formatCount; f++) {
                    DWORD fmt = 0, size = 0;
                    if (!readDword(fmt) || !readDword(size)) { ok = false; break; }
                    if (off + size > len) { ok = false; break; }
                    if (size > 0 && size <= kPersistFormatCap)
                        fmts[fmt] = std::vector<BYTE>(p + off, p + off + size);
                    off += size;
                }
                if (!ok) break;
                auto item = MakeItemFromFormats(fmts, pinned != 0);
                if (item) clipboardHistory.push_back(std::move(item));
            }
        }
        if (!clipboardHistory.empty()) FilterItems();
    } catch (...) {}
}

void ClipboardManager::SaveClipboardHistory() {
    try {
        wchar_t appData[MAX_PATH];
        if (FAILED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, 0, appData))) return;
        std::wstring path = std::wstring(appData) + L"\\clip2";
        CreateDirectoryW(path.c_str(), nullptr);
        path += L"\\history.dat";

        // Decide which items to save: all pinned first (exempt from the count cap),
        // then the most recent unpinned items up to maxItems.
        size_t cap = (size_t)std::max(MIN_MAX_ITEMS, std::min(maxItems, MAX_MAX_ITEMS));
        std::vector<size_t> order;
        for (size_t i = 0; i < clipboardHistory.size(); i++)
            if (clipboardHistory[i] && clipboardHistory[i]->pinned) order.push_back(i);
        for (size_t i = 0; i < clipboardHistory.size() && order.size() < cap + 0; i++) {
            if (clipboardHistory[i] && !clipboardHistory[i]->pinned) {
                if (order.size() >= cap) break;
                order.push_back(i);
            }
        }

        // Build the inner plaintext payload.
        std::vector<BYTE> payload;
        AppendBytes(payload, "CLP2", 4);
        AppendDword(payload, 2);  // version 2 (all formats + pinned)
        size_t countPos = payload.size();
        AppendDword(payload, 0);  // placeholder for count
        DWORD savedCount = 0;

        for (size_t idx : order) {
            const auto& item = clipboardHistory[idx];
            if (!item) continue;
            // Collect formats that fit within the per-format and per-item caps.
            std::vector<std::pair<UINT, const std::vector<BYTE>*>> keep;
            size_t itemTotal = 0;
            for (const auto& kv : item->formats) {
                if (kv.second.empty() || kv.second.size() > kPersistFormatCap) continue;
                if (itemTotal + kv.second.size() > kPersistItemCap) continue;
                keep.emplace_back(kv.first, &kv.second);
                itemTotal += kv.second.size();
            }
            if (keep.empty()) continue;
            payload.push_back(item->pinned ? 1 : 0);
            AppendDword(payload, (DWORD)keep.size());
            for (const auto& kv : keep) {
                AppendDword(payload, (DWORD)kv.first);
                AppendDword(payload, (DWORD)kv.second->size());
                AppendBytes(payload, kv.second->data(), kv.second->size());
            }
            savedCount++;
        }
        memcpy(payload.data() + countPos, &savedCount, sizeof(DWORD));

        // Encrypt at rest with DPAPI; fall back to plaintext only if encryption fails.
        HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;
        DWORD written = 0;
        std::vector<BYTE> blob;
        if (DpapiProtect(payload, blob)) {
            DWORD blobLen = (DWORD)blob.size();
            WriteFile(h, "CLP3", 4, &written, nullptr);
            WriteFile(h, &blobLen, 4, &written, nullptr);
            WriteFile(h, blob.data(), blobLen, &written, nullptr);
        } else {
            WriteFile(h, payload.data(), (DWORD)payload.size(), &written, nullptr);
        }
        CloseHandle(h);
    } catch (...) {}
}

void ClipboardManager::FilterItems() {
    filteredIndices.clear();
    ClearMultiSelection(); // Clear multi-selection when filtering
    
    if (searchText.empty()) {
        // No filter - show all items in chronological order.
        for (size_t i = 0; i < clipboardHistory.size(); i++) {
            filteredIndices.push_back((int)i);
        }
    } else {
        // Fuzzy filter + ranking (case-insensitive). Best matches first.
        std::wstring searchLower = searchText;
        std::transform(searchLower.begin(), searchLower.end(), searchLower.begin(), ::towlower);
        
        struct Scored { int index; int score; int order; };
        std::vector<Scored> scored;
        for (size_t i = 0; i < clipboardHistory.size(); i++) {
            const auto& item = clipboardHistory[i];
            
            // Score full text content (handles line breaks, not just first 50 chars).
            int best = FuzzyScore(searchLower, item->GetFullSearchableText());
            // Also consider preview / format name / file type so images and files are findable.
            best = std::max(best, FuzzyScore(searchLower, item->preview));
            best = std::max(best, FuzzyScore(searchLower, item->formatName));
            best = std::max(best, FuzzyScore(searchLower, item->fileType));
            
            if (best > 0) {
                scored.push_back({ (int)i, best, (int)i });
            }
        }
        // Stable sort: higher score first, then original (chronological) order.
        std::stable_sort(scored.begin(), scored.end(), [](const Scored& a, const Scored& b) {
            if (a.score != b.score) return a.score > b.score;
            return a.order < b.order;
        });
        for (const auto& s : scored) filteredIndices.push_back(s.index);
    }

    // Pinned items always float to the top, preserving their relative order.
    std::stable_partition(filteredIndices.begin(), filteredIndices.end(), [this](int idx) {
        return idx >= 0 && idx < (int)clipboardHistory.size() &&
               clipboardHistory[idx] && clipboardHistory[idx]->pinned;
    });
    
    // Reset scroll offset and selection when filtering
    scrollOffset = 0;
    if (filteredIndices.empty()) {
        selectedIndex = -1;
    } else {
        selectedIndex = 0;
    }
    UpdateListWindow();
}

void ClipboardManager::FilterSnippets() {
    filteredSnippetIndices.clear();
    
    if (searchText.empty()) {
        for (size_t i = 0; i < snippets.size(); i++) {
            filteredSnippetIndices.push_back((int)i);
        }
    } else {
        std::wstring searchLower = searchText;
        std::transform(searchLower.begin(), searchLower.end(), searchLower.begin(), ::towlower);
        
        struct Scored { int index; int score; int order; };
        std::vector<Scored> scored;
        for (size_t i = 0; i < snippets.size(); i++) {
            int best = FuzzyScore(searchLower, snippets[i].name);
            best = std::max(best, FuzzyScore(searchLower, snippets[i].content));
            if (!snippets[i].contentPlain.empty())
                best = std::max(best, FuzzyScore(searchLower, snippets[i].contentPlain));
            if (best > 0) scored.push_back({ (int)i, best, (int)i });
        }
        std::stable_sort(scored.begin(), scored.end(), [](const Scored& a, const Scored& b) {
            if (a.score != b.score) return a.score > b.score;
            return a.order < b.order;
        });
        for (const auto& s : scored) filteredSnippetIndices.push_back(s.index);
    }
    
    scrollOffset = 0;
    if (filteredSnippetIndices.empty()) {
        selectedIndex = -1;
    } else {
        selectedIndex = 0;
    }
    UpdateListWindow();
}

void ClipboardManager::SetStartupWithWindows(bool enable) {
    HKEY hKey;
    LONG result = RegOpenKeyExW(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0,
        KEY_SET_VALUE,
        &hKey
    );
    
    if (result == ERROR_SUCCESS) {
        if (enable) {
            // Get executable path
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(GetModuleHandle(nullptr), exePath, MAX_PATH);
            
            // Add to startup
            RegSetValueExW(
                hKey,
                L"clip2",
                0,
                REG_SZ,
                (BYTE*)exePath,
                (DWORD)(wcslen(exePath) + 1) * sizeof(wchar_t)
            );
        } else {
            // Remove from startup
            RegDeleteValueW(hKey, L"clip2");
        }
        
        RegCloseKey(hKey);
    }
}

bool ClipboardManager::IsStartupWithWindows() {
    HKEY hKey;
    LONG result = RegOpenKeyExW(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0,
        KEY_READ,
        &hKey
    );
    
    if (result == ERROR_SUCCESS) {
        wchar_t exePath[MAX_PATH];
        DWORD dataSize = sizeof(exePath);
        DWORD type = REG_SZ;
        
        result = RegQueryValueExW(
            hKey,
            L"clip2",
            nullptr,
            &type,
            (LPBYTE)exePath,
            &dataSize
        );
        
        RegCloseKey(hKey);
        
        if (result == ERROR_SUCCESS) {
            // Verify it's our executable
            wchar_t currentExePath[MAX_PATH];
            GetModuleFileNameW(GetModuleHandle(nullptr), currentExePath, MAX_PATH);
            
            // Compare paths (case-insensitive)
            return _wcsicmp(exePath, currentExePath) == 0;
        }
    }
    
    return false;
}

static void RegReadHotkey(HKEY hKey, const wchar_t* modName, const wchar_t* vkName, ClipboardManager::HotkeyConfig& hk) {
    DWORD val = 0, sz = sizeof(DWORD), type = REG_DWORD;
    if (RegQueryValueExW(hKey, modName, nullptr, &type, (LPBYTE)&val, &sz) == ERROR_SUCCESS && type == REG_DWORD)
        hk.modifiers = val;
    sz = sizeof(DWORD); type = REG_DWORD;
    if (RegQueryValueExW(hKey, vkName, nullptr, &type, (LPBYTE)&val, &sz) == ERROR_SUCCESS && type == REG_DWORD)
        hk.vkCode = val;
}

static void RegWriteHotkey(HKEY hKey, const wchar_t* modName, const wchar_t* vkName, const ClipboardManager::HotkeyConfig& hk) {
    DWORD m = hk.modifiers, v = hk.vkCode;
    RegSetValueExW(hKey, modName, 0, REG_DWORD, (BYTE*)&m, sizeof(DWORD));
    RegSetValueExW(hKey, vkName, 0, REG_DWORD, (BYTE*)&v, sizeof(DWORD));
}

void ClipboardManager::LoadHotkeyConfig() {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return;
    RegReadHotkey(hKey, L"HotkeyModifiers",          L"HotkeyVkCode",          hotkeyConfig);
    RegReadHotkey(hKey, L"SnippetsModifiers",         L"SnippetsVkCode",        snippetsHotkey);
    RegReadHotkey(hKey, L"CopyFocusedModifiers",      L"CopyFocusedVkCode",     copyFocusedHotkey);
    RegReadHotkey(hKey, L"PasteFocusedModifiers",     L"PasteFocusedVkCode",    pasteFocusedHotkey);
    RegReadHotkey(hKey, L"PasteClipboardModifiers",   L"PasteClipboardVkCode",  pasteClipboardHotkey);
    RegCloseKey(hKey);

    // Migration: the legacy default snippet hotkey was Ctrl+NumPad0. It has been removed,
    // so users with that combo persisted from an older build get it cleared automatically.
    if (snippetsHotkey.modifiers == MOD_CONTROL && snippetsHotkey.vkCode == VK_NUMPAD0) {
        snippetsHotkey.modifiers = 0;
        snippetsHotkey.vkCode = 0;
    }
}

void ClipboardManager::SaveHotkeyConfig() {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2", 0, nullptr,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr) != ERROR_SUCCESS)
        return;
    RegWriteHotkey(hKey, L"HotkeyModifiers",          L"HotkeyVkCode",          hotkeyConfig);
    RegWriteHotkey(hKey, L"SnippetsModifiers",         L"SnippetsVkCode",        snippetsHotkey);
    RegWriteHotkey(hKey, L"CopyFocusedModifiers",      L"CopyFocusedVkCode",     copyFocusedHotkey);
    RegWriteHotkey(hKey, L"PasteFocusedModifiers",     L"PasteFocusedVkCode",    pasteFocusedHotkey);
    RegWriteHotkey(hKey, L"PasteClipboardModifiers",   L"PasteClipboardVkCode",  pasteClipboardHotkey);
    RegCloseKey(hKey);
}

// Update the class background brushes so newly-erased regions use the active BG, then
// force a full repaint of everything that consumes Theme5250::* in WM_PAINT.
void ClipboardManager::RefreshThemeVisuals() {
    auto refreshClassBrush = [](HWND hwnd, COLORREF c) {
        if (!hwnd) return;
        HBRUSH oldBrush = (HBRUSH)GetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND);
        HBRUSH newBrush = CreateSolidBrush(c);
        if (newBrush) {
            SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, (LONG_PTR)newBrush);
            if (oldBrush) DeleteObject(oldBrush);
        }
    };
    refreshClassBrush(hwndList, Theme5250::BG);
    refreshClassBrush(hwndPreview, Theme5250::BG);

    if (hwndList) { InvalidateRect(hwndList, nullptr, TRUE); UpdateWindow(hwndList); }
    if (hwndPreview && IsWindowVisible(hwndPreview)) {
        InvalidateRect(hwndPreview, nullptr, TRUE); UpdateWindow(hwndPreview);
    }
    if (hwndSearch) { InvalidateRect(hwndSearch, nullptr, TRUE); UpdateWindow(hwndSearch); }
}

void ClipboardManager::SetTheme(int themeId) {
    ApplyThemeId(themeId);
    SaveThemeIdToRegistry(themeId);
    RefreshThemeVisuals();
}

void ClipboardManager::SetThemeFontColor(COLORREF color) {
    g_themeFontColor = color;
    SaveThemeFontColorToRegistry(color);
    // ApplyThemeId re-reads g_themeFontColor and updates TXT/SEL_BG accordingly.
    ApplyThemeId(g_currentThemeId);
    RefreshThemeVisuals();
}

void ClipboardManager::SetThemeColorOverride(int slot, COLORREF color) {
    if (slot < 0 || slot >= TC_COUNT) return;
    g_colorOverride[slot] = color;
    SaveThemeColorToRegistry(slot, color);
    ApplyThemeId(g_currentThemeId);
    RefreshThemeVisuals();
}

void ClipboardManager::ResetThemeColorOverrides() {
    for (int i = 0; i < TC_COUNT; i++) {
        g_colorOverride[i] = kColorUsePreset;
        SaveThemeColorToRegistry(i, kColorUsePreset);
    }
    ApplyThemeId(g_currentThemeId);
    RefreshThemeVisuals();
}

void ClipboardManager::SetThemeFontFace(const std::wstring& face) {
    std::wstring resolved = face;
    if (resolved.empty() || !FontFamilyExists(resolved))
        resolved = kDefaultThemeFontFace;
    g_themeFontFace = resolved;
    SaveThemeFontFaceToRegistry(resolved);
    // Drop the cached HFONT so the next paint and the WM_SETFONT below pick up the new face.
    ResetOverlayFontCache();
    HFONT newFont = GetOverlayFont();
    if (hwndSearch && newFont) SendMessage(hwndSearch, WM_SETFONT, (WPARAM)newFont, TRUE);
    if (hwndList) { InvalidateRect(hwndList, nullptr, TRUE); UpdateWindow(hwndList); }
    if (hwndPreview && IsWindowVisible(hwndPreview)) {
        InvalidateRect(hwndPreview, nullptr, TRUE); UpdateWindow(hwndPreview);
    }
    if (hwndSearch) { InvalidateRect(hwndSearch, nullptr, TRUE); UpdateWindow(hwndSearch); }
}

void ClipboardManager::LoadSnippets() {
    snippets.clear();
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\clip2\\Snippets", 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) return;

    DWORD count = 0;
    DWORD dataSize = sizeof(DWORD);
    DWORD type = REG_DWORD;
    result = RegQueryValueExW(hKey, L"Count", nullptr, &type, (LPBYTE)&count, &dataSize);
    if (result != ERROR_SUCCESS || type != REG_DWORD) { RegCloseKey(hKey); return; }

    wchar_t valueName[32];
    for (DWORD i = 0; i < count && i < 500; i++) {
        swprintf_s(valueName, L"N%d", i);
        wchar_t buf[65536];  // 32KB max per snippet
        dataSize = sizeof(buf);
        type = REG_SZ;
        result = RegQueryValueExW(hKey, valueName, nullptr, &type, (LPBYTE)buf, &dataSize);
        if (result == ERROR_SUCCESS && type == REG_SZ) {
            std::wstring full(buf);
            size_t sep1 = full.find(L'\x01');
            if (sep1 != std::wstring::npos) {
                Snippet s;
                s.name = full.substr(0, sep1);
                size_t sep2 = full.find(L'\x02', sep1 + 1);
                if (sep2 != std::wstring::npos) {
                    s.content = full.substr(sep1 + 1, sep2 - sep1 - 1);
                    s.contentPlain = full.substr(sep2 + 1);
                } else {
                    s.content = full.substr(sep1 + 1);
                }
                snippets.push_back(std::move(s));
            }
        }
    }
    RegCloseKey(hKey);
}

void ClipboardManager::SaveSnippets() {
    HKEY hKey;
    LONG result = RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\clip2\\Snippets", 0, nullptr,
        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);
    if (result != ERROR_SUCCESS) return;

    DWORD count = (DWORD)snippets.size();
    RegSetValueExW(hKey, L"Count", 0, REG_DWORD, (BYTE*)&count, sizeof(DWORD));

    wchar_t valueName[32];
    for (size_t i = 0; i < snippets.size(); i++) {
        swprintf_s(valueName, L"N%zu", i);
        std::wstring full = snippets[i].name + L'\x01' + snippets[i].content;
        if (!snippets[i].contentPlain.empty()) full += L'\x02' + snippets[i].contentPlain;
        if (full.size() < 32767) {
            RegSetValueExW(hKey, valueName, 0, REG_SZ, (BYTE*)full.c_str(),
                (DWORD)(full.size() + 1) * sizeof(wchar_t));
        }
    }
    RegCloseKey(hKey);
}

std::wstring ClipboardManager::ExpandSnippetPlaceholders(const std::wstring& content) {
    std::wstring result = content;
    auto now = std::chrono::system_clock::now();
    time_t t = std::chrono::system_clock::to_time_t(now);
    struct tm tmBuf;
    if (localtime_s(&tmBuf, &t) == 0) {
        wchar_t dateBuf[64], timeBuf[64], datetimeBuf[128];
        wcsftime(dateBuf, 64, L"%Y-%m-%d", &tmBuf);
        wcsftime(timeBuf, 64, L"%H:%M:%S", &tmBuf);
        wcsftime(datetimeBuf, 128, L"%Y-%m-%d %H:%M:%S", &tmBuf);

        auto replace = [&result](const wchar_t* from, const std::wstring& to) {
            size_t pos = 0;
            while ((pos = result.find(from, pos)) != std::wstring::npos) {
                result.replace(pos, wcslen(from), to);
                pos += to.size();
            }
        };
        replace(L"{{date}}", dateBuf);
        replace(L"{{time}}", timeBuf);
        replace(L"{{datetime}}", datetimeBuf);
        replace(L"{{year}}", std::to_wstring(tmBuf.tm_year + 1900));
        replace(L"{{month}}", std::to_wstring(tmBuf.tm_mon + 1));
        replace(L"{{day}}", std::to_wstring(tmBuf.tm_mday));
        replace(L"{{hour}}", std::to_wstring(tmBuf.tm_hour));
        replace(L"{{minute}}", std::to_wstring(tmBuf.tm_min));
        replace(L"{{second}}", std::to_wstring(tmBuf.tm_sec));
    }

    if (result.find(L"{{clipboard}}") != std::wstring::npos) {
        std::wstring clipText;
        if (OpenClipboard(hwndMain)) {
            HANDLE h = GetClipboardData(CF_UNICODETEXT);
            if (h) {
                const wchar_t* p = (const wchar_t*)GlobalLock(h);
                if (p) {
                    clipText = p;
                    GlobalUnlock(h);
                }
            }
            CloseClipboard();
        }
        size_t pos = 0;
        while ((pos = result.find(L"{{clipboard}}", pos)) != std::wstring::npos) {
            result.replace(pos, 13, clipText);
            pos += clipText.size();
        }
    }
    return result;
}

void ClipboardManager::PasteSnippet(int index) {
    if (index < 0 || index >= (int)snippets.size()) return;
    const Snippet& snip = snippets[index];

    std::wstring storedContent = ExpandSnippetPlaceholders(snip.content);
    bool isRtf = IsRtfContent(storedContent);

    std::string rtfBytes;
    std::wstring plainText;
    if (isRtf) {
        // Rich snippet: stored byte-by-byte as a wide string. Recover the bytes; never
        // expose those bytes as CF_UNICODETEXT — that's the bug that made snippets paste
        // as RTF source text instead of formatted content.
        rtfBytes = RtfBytesFromStoredWide(storedContent);
        if (!snip.contentPlain.empty()) {
            plainText = ExpandSnippetPlaceholders(snip.contentPlain);
        } else {
            // Legacy snippets without contentPlain: derive plain text from the RTF itself.
            plainText = ExtractPlainFromRtfBytes(rtfBytes);
        }
    } else {
        plainText = storedContent;
    }

    if (rtfBytes.empty() && plainText.empty()) return;

    isPasting = true;

    if (!OpenClipboardWithRetry(hwndMain, 24, 8)) {
        isPasting = false;
        return;
    }
    EmptyClipboard();

    UINT cfRtf = CfRtf();
    bool rtfSet = false;
    bool plainSet = false;
    if (isRtf && cfRtf != 0 && !rtfBytes.empty()) {
        rtfSet = PutBytesOnClipboard(cfRtf, rtfBytes.data(), rtfBytes.size() + 1);
    }
    if (!plainText.empty()) {
        size_t bytes = (plainText.size() + 1) * sizeof(wchar_t);
        plainSet = PutBytesOnClipboard(CF_UNICODETEXT, plainText.c_str(), bytes);
    }
    CloseClipboard();

    if (!rtfSet && !plainSet) {
        isPasting = false;
        return;
    }

    lastPastedText = plainText;
    lastSequenceNumber = GetClipboardSequenceNumber();

    if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
        SetForegroundWindow(previousFocusWindow);
        SetFocus(previousFocusWindow);
        Sleep(40);
    }
    ReleaseHotkeyModifiersForPaste();
    Sleep(30);
    SendCtrlV();
    Sleep(460);
    isPasting = false;
    PlayClickSound();
}

void ClipboardManager::ShowSnippetsManagerDialog() {
    if (hwndSnippetsManager && IsWindow(hwndSnippetsManager)) {
        SetForegroundWindow(hwndSnippetsManager);
        return;
    }
    static bool msfteditLoaded = false;
    if (!msfteditLoaded) {
        LoadLibraryW(L"Msftedit.dll");
        msfteditLoaded = true;
    }
    WNDCLASS wc = {};
    wc.lpfnWndProc = SnippetsManagerProc;
    wc.hInstance = GetModuleHandle(nullptr);
    wc.lpszClassName = L"Clip2SnippetsManagerClass";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    RegisterClass(&wc);

    hwndSnippetsManager = CreateWindowExW(0, L"Clip2SnippetsManagerClass", L"clip2 - Snippets",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        100, 100, 450, 400, hwndMain, nullptr, GetModuleHandle(nullptr), nullptr);

    if (!hwndSnippetsManager) return;

    CreateWindowExW(0, L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_BORDER,
        10, 10, 150, 280, hwndSnippetsManager, (HMENU)1001, GetModuleHandle(nullptr), nullptr);

    CreateWindowExW(0, L"STATIC", L"Name:", WS_CHILD | WS_VISIBLE, 170, 10, 260, 18,
        hwndSnippetsManager, nullptr, GetModuleHandle(nullptr), nullptr);
    CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER, 170, 28, 260, 22,
        hwndSnippetsManager, (HMENU)1002, GetModuleHandle(nullptr), nullptr);

    CreateWindowExW(0, L"STATIC", L"Content:", WS_CHILD | WS_VISIBLE, 170, 58, 260, 18,
        hwndSnippetsManager, nullptr, GetModuleHandle(nullptr), nullptr);
    CreateWindowExW(WS_EX_CLIENTEDGE, MSFTEDIT_CLASS, L"",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_NOOLEDRAGDROP,
        170, 76, 260, 214, hwndSnippetsManager, (HMENU)1003, GetModuleHandle(nullptr), nullptr);

    CreateWindowExW(0, L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        10, 300, 70, 28, hwndSnippetsManager, (HMENU)1004, GetModuleHandle(nullptr), nullptr);
    CreateWindowExW(0, L"BUTTON", L"Update", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        85, 300, 70, 28, hwndSnippetsManager, (HMENU)1005, GetModuleHandle(nullptr), nullptr);
    CreateWindowExW(0, L"BUTTON", L"Delete", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        160, 300, 70, 28, hwndSnippetsManager, (HMENU)1006, GetModuleHandle(nullptr), nullptr);
    CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        350, 300, 70, 28, hwndSnippetsManager, (HMENU)1007, GetModuleHandle(nullptr), nullptr);

    HWND hList = GetDlgItem(hwndSnippetsManager, 1001);
    for (const auto& s : snippets) {
        SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)s.name.c_str());
    }

    static const wchar_t* placeholderHelp = L"Placeholders: {{date}}, {{time}}, {{datetime}}, {{year}}, {{month}}, {{day}}, {{hour}}, {{minute}}, {{second}}, {{clipboard}}";
    CreateWindowExW(0, L"STATIC", placeholderHelp, WS_CHILD | WS_VISIBLE | SS_CENTER,
        10, 335, 410, 30, hwndSnippetsManager, nullptr, GetModuleHandle(nullptr), nullptr);

    ShowWindow(hwndSnippetsManager, SW_SHOW);
}

LRESULT CALLBACK ClipboardManager::SnippetsManagerProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    static int selectedIdx = -1;

    switch (uMsg) {
    case WM_CLOSE:
        if (mgr && mgr->hwndSnippetsManager == hwnd) mgr->hwndSnippetsManager = nullptr;
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        if (mgr && mgr->hwndSnippetsManager == hwnd) mgr->hwndSnippetsManager = nullptr;
        return 0;

    case WM_COMMAND:
        if (!mgr) break;
        if (LOWORD(wParam) == 1001 && HIWORD(wParam) == LBN_SELCHANGE) {
            HWND hList = GetDlgItem(hwnd, 1001);
            HWND hContent = GetDlgItem(hwnd, 1003);
            int i = (int)SendMessageW(hList, LB_GETCURSEL, 0, 0);
            if (i >= 0 && i < (int)mgr->snippets.size()) {
                selectedIdx = i;
                SetDlgItemTextW(hwnd, 1002, mgr->snippets[i].name.c_str());
                if (IsRtfContent(mgr->snippets[i].content)) {
                    std::string rtfBytes;
                    for (wchar_t wc : mgr->snippets[i].content) rtfBytes += (char)(wc & 0xFF);
                    StreamCookie cookie = { rtfBytes.c_str(), 0, rtfBytes.size(), nullptr };
                    EDITSTREAM es = { (DWORD_PTR)&cookie, 0, RichEditStreamInCallback };
                    SendMessage(hContent, EM_STREAMIN, SF_RTF, (LPARAM)&es);
                } else {
                    SetDlgItemTextW(hwnd, 1003, mgr->snippets[i].content.c_str());
                }
            }
        } else if (LOWORD(wParam) == 1004) {  // Add
            wchar_t nameBuf[256];
            GetDlgItemTextW(hwnd, 1002, nameBuf, 256);
            if (nameBuf[0]) {
                HWND hContent = GetDlgItem(hwnd, 1003);
                std::string rtfOut;
                StreamCookie cookie = { nullptr, 0, 0, &rtfOut };
                EDITSTREAM es = { (DWORD_PTR)&cookie, 0, RichEditStreamOutCallback };
                SendMessage(hContent, EM_STREAMOUT, SF_RTF, (LPARAM)&es);
                std::wstring content, contentPlain;
                if (!rtfOut.empty() && rtfOut.size() >= 5 && rtfOut[0] == '{' && rtfOut[1] == '\\' && (rtfOut[2] == 'r' || rtfOut[2] == 'R')) {
                    content.resize(rtfOut.size());
                    for (size_t j = 0; j < rtfOut.size(); j++) content[j] = (wchar_t)(unsigned char)rtfOut[j];
                    wchar_t plainBuf[32768];
                    GetWindowTextW(hContent, plainBuf, 32768);
                    contentPlain = plainBuf;
                } else {
                    wchar_t contentBuf[32768];
                    GetWindowTextW(hContent, contentBuf, 32768);
                    content = contentBuf;
                }
                if (_wcsicmp(nameBuf, L"*set") == 0) {
                    MessageBoxW(hwnd, L"Cannot use '*set' as snippet name (reserved).", L"clip2 Snippets", MB_OK | MB_ICONWARNING);
                } else {
                    Snippet s;
                    s.name = nameBuf;
                    s.content = content;
                    s.contentPlain = contentPlain;
                    mgr->snippets.push_back(s);
                    mgr->SaveSnippets();
                    SendMessageW(GetDlgItem(hwnd, 1001), LB_ADDSTRING, 0, (LPARAM)s.name.c_str());
                    SetDlgItemTextW(hwnd, 1002, L"");
                    SendMessage(GetDlgItem(hwnd, 1003), WM_SETTEXT, 0, (LPARAM)L"");
                    selectedIdx = -1;
                }
            }
        } else if (LOWORD(wParam) == 1005 && selectedIdx >= 0) {  // Update
            wchar_t nameBuf[256];
            GetDlgItemTextW(hwnd, 1002, nameBuf, 256);
            if (nameBuf[0] && selectedIdx < (int)mgr->snippets.size()) {
                HWND hContent = GetDlgItem(hwnd, 1003);
                std::string rtfOut;
                StreamCookie cookie = { nullptr, 0, 0, &rtfOut };
                EDITSTREAM es = { (DWORD_PTR)&cookie, 0, RichEditStreamOutCallback };
                SendMessage(hContent, EM_STREAMOUT, SF_RTF, (LPARAM)&es);
                std::wstring content, contentPlain;
                if (!rtfOut.empty() && rtfOut.size() >= 5 && rtfOut[0] == '{' && rtfOut[1] == '\\' && (rtfOut[2] == 'r' || rtfOut[2] == 'R')) {
                    content.resize(rtfOut.size());
                    for (size_t j = 0; j < rtfOut.size(); j++) content[j] = (wchar_t)(unsigned char)rtfOut[j];
                    wchar_t plainBuf[32768];
                    GetWindowTextW(hContent, plainBuf, 32768);
                    contentPlain = plainBuf;
                } else {
                    wchar_t contentBuf[32768];
                    GetWindowTextW(hContent, contentBuf, 32768);
                    content = contentBuf;
                }
                if (_wcsicmp(nameBuf, L"*set") == 0) {
                    MessageBoxW(hwnd, L"Cannot use '*set' as snippet name (reserved).", L"clip2 Snippets", MB_OK | MB_ICONWARNING);
                } else {
                    mgr->snippets[selectedIdx].name = nameBuf;
                    mgr->snippets[selectedIdx].content = content;
                    mgr->snippets[selectedIdx].contentPlain = contentPlain;
                    mgr->SaveSnippets();
                    HWND hList = GetDlgItem(hwnd, 1001);
                    SendMessageW(hList, LB_DELETESTRING, selectedIdx, 0);
                    SendMessageW(hList, LB_INSERTSTRING, selectedIdx, (LPARAM)nameBuf);
                    SendMessageW(hList, LB_SETCURSEL, selectedIdx, 0);
                }
            }
        } else if (LOWORD(wParam) == 1006 && selectedIdx >= 0) {  // Delete
            if (selectedIdx < (int)mgr->snippets.size()) {
                mgr->snippets.erase(mgr->snippets.begin() + selectedIdx);
                mgr->SaveSnippets();
                SendMessageW(GetDlgItem(hwnd, 1001), LB_DELETESTRING, selectedIdx, 0);
                SetDlgItemTextW(hwnd, 1002, L"");
                SetDlgItemTextW(hwnd, 1003, L"");
                selectedIdx = (selectedIdx < (int)mgr->snippets.size()) ? selectedIdx : (int)mgr->snippets.size() - 1;
                if (selectedIdx >= 0) {
                    SendMessageW(GetDlgItem(hwnd, 1001), LB_SETCURSEL, selectedIdx, 0);
                    SetDlgItemTextW(hwnd, 1002, mgr->snippets[selectedIdx].name.c_str());
                    HWND hContent = GetDlgItem(hwnd, 1003);
                    if (IsRtfContent(mgr->snippets[selectedIdx].content)) {
                        std::string rtfBytes;
                        for (wchar_t wc : mgr->snippets[selectedIdx].content) rtfBytes += (char)(wc & 0xFF);
                        StreamCookie ck = { rtfBytes.c_str(), 0, rtfBytes.size(), nullptr };
                        EDITSTREAM eds = { (DWORD_PTR)&ck, 0, RichEditStreamInCallback };
                        SendMessage(hContent, EM_STREAMIN, SF_RTF, (LPARAM)&eds);
                    } else {
                        SetWindowTextW(hContent, mgr->snippets[selectedIdx].content.c_str());
                    }
                }
            }
        } else if (LOWORD(wParam) == 1007) {  // Close
            PostMessage(hwnd, WM_CLOSE, 0, 0);
        }
        break;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

std::wstring ClipboardManager::HotkeyToString(const HotkeyConfig& hk) {
    if (hk.vkCode == 0) return L"(none)";
    std::wstring s;
    if (hk.modifiers & MOD_CONTROL) s += L"Ctrl+";
    if (hk.modifiers & MOD_ALT)     s += L"Alt+";
    if (hk.modifiers & MOD_SHIFT)   s += L"Shift+";
    if (hk.modifiers & MOD_WIN)     s += L"Win+";
    wchar_t keyName[128] = {0};
    UINT sc = MapVirtualKeyW(hk.vkCode, MAPVK_VK_TO_VSC);
    if (sc) {
        LONG lp = (sc << 16);
        if (hk.vkCode >= VK_NUMPAD0 && hk.vkCode <= VK_NUMPAD9) lp |= (1 << 24);
        if (hk.vkCode == VK_DECIMAL || hk.vkCode == VK_INSERT || hk.vkCode == VK_DELETE ||
            hk.vkCode == VK_HOME || hk.vkCode == VK_END || hk.vkCode == VK_PRIOR ||
            hk.vkCode == VK_NEXT || hk.vkCode == VK_UP || hk.vkCode == VK_DOWN ||
            hk.vkCode == VK_LEFT || hk.vkCode == VK_RIGHT) lp |= (1 << 24);
        if (GetKeyNameTextW(lp, keyName, 128))
            s += keyName;
        else
            s += L"0x" + std::to_wstring(hk.vkCode);
    } else {
        s += L"0x" + std::to_wstring(hk.vkCode);
    }
    return s;
}

bool ClipboardManager::HotkeyFromKeyMessage(UINT vk, HotkeyConfig& out) {
    if (vk == VK_CONTROL || vk == VK_LCONTROL || vk == VK_RCONTROL ||
        vk == VK_MENU || vk == VK_LMENU || vk == VK_RMENU ||
        vk == VK_SHIFT || vk == VK_LSHIFT || vk == VK_RSHIFT ||
        vk == VK_LWIN || vk == VK_RWIN)
        return false;
    UINT mod = 0;
    if (GetAsyncKeyState(VK_CONTROL) & 0x8000) mod |= MOD_CONTROL;
    if (GetAsyncKeyState(VK_MENU)    & 0x8000) mod |= MOD_ALT;
    if (GetAsyncKeyState(VK_SHIFT)   & 0x8000) mod |= MOD_SHIFT;
    if ((GetAsyncKeyState(VK_LWIN) | GetAsyncKeyState(VK_RWIN)) & 0x8000) mod |= MOD_WIN;
    out.modifiers = mod;
    out.vkCode = vk;
    return true;
}

static const int IDC_HK_OVERLAY          = 3001;
static const int IDC_HK_SNIPPETS         = 3002;
static const int IDC_HK_COPY_FOCUSED     = 3003;
static const int IDC_HK_PASTE_FOCUSED    = 3004;
static const int IDC_HK_PASTE_CLIPBOARD  = 3005;
static const int IDC_THEME_COMBO         = 3006;
static const int IDC_FONT_COMBO          = 3007;
static const int IDC_FONT_COLOR_BTN      = 3008;
static const int IDC_FONT_COLOR_SWATCH   = 3009;
static const int IDC_BTN_SAVE            = 3010;
static const int IDC_BTN_CANCEL          = 3011;
static const int IDC_BTN_DEFAULTS        = 3012;
static const int IDC_HISTORY_SIZE        = 3013;
static const int IDC_COLOR_SWATCH_FIRST  = 3014;  // 6 element-color swatches: [FIRST, FIRST+TC_COUNT)

static HWND CreateHotkeyLabel(HWND parent, const wchar_t* text, int x, int y, int w, int h, HFONT font) {
    HWND lbl = CreateWindowW(L"STATIC", text, WS_CHILD | WS_VISIBLE | SS_LEFT,
                             x, y, w, h, parent, nullptr, GetModuleHandle(nullptr), nullptr);
    SendMessage(lbl, WM_SETFONT, (WPARAM)font, TRUE);
    return lbl;
}

static HWND CreateHotkeyDisplay(HWND parent, int id, const std::wstring& text, int x, int y, int w, int h, HFONT font) {
    HWND ctrl = CreateWindowW(L"EDIT", text.c_str(),
        WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_CENTER | ES_READONLY,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(nullptr), nullptr);
    SendMessage(ctrl, WM_SETFONT, (WPARAM)font, TRUE);
    return ctrl;
}

LRESULT CALLBACK ClipboardManager::SettingsHotkeyEditSubclass(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    ClipboardManager* self = reinterpret_cast<ClipboardManager*>(dwRefData);
    if (!self) return DefSubclassProc(hWnd, uMsg, wParam, lParam);

    if (uMsg == WM_KEYDOWN || uMsg == WM_SYSKEYDOWN) {
        if (wParam == VK_TAB)
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        HotkeyConfig captured;
        if (HotkeyFromKeyMessage(static_cast<UINT>(wParam), captured)) {
            const int id = GetDlgCtrlID(hWnd);
            HotkeyConfig* target = nullptr;
            if (id == IDC_HK_OVERLAY) target = &self->hotkeyConfig;
            else if (id == IDC_HK_SNIPPETS) target = &self->snippetsHotkey;
            else if (id == IDC_HK_COPY_FOCUSED) target = &self->copyFocusedHotkey;
            else if (id == IDC_HK_PASTE_FOCUSED) target = &self->pasteFocusedHotkey;
            else if (id == IDC_HK_PASTE_CLIPBOARD) target = &self->pasteClipboardHotkey;
            if (target) {
                *target = captured;
                SetWindowTextW(hWnd, HotkeyToString(captured).c_str());
            }
            return 0;
        }
        return 0;
    }

    if (uMsg == WM_CHAR || uMsg == WM_SYSCHAR) {
        return 0;
    }

    if (uMsg == WM_NCDESTROY)
        RemoveWindowSubclass(hWnd, SettingsHotkeyEditSubclass, uIdSubclass);

    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void ClipboardManager::ShowSettingsDialog() {
    if (hwndSettings && IsWindow(hwndSettings)) {
        SetForegroundWindow(hwndSettings);
        return;
    }

    const int DLG_W = 480, DLG_H = 620;
    RECT workArea;
    SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
    int x = workArea.left + ((workArea.right - workArea.left) - DLG_W) / 2;
    int y = workArea.top  + ((workArea.bottom - workArea.top) - DLG_H) / 2;

    WNDCLASSW wc = {};
    wc.lpfnWndProc = SettingsDialogProc;
    wc.hInstance = GetModuleHandle(nullptr);
    wc.lpszClassName = L"clip2_SettingsDialog";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    RegisterClassW(&wc);

    hwndSettings = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"clip2_SettingsDialog", L"clip2 Settings \u2014 Shortcuts & Theme",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        x, y, DLG_W, DLG_H, hwndMain, nullptr, GetModuleHandle(nullptr), nullptr);

    HFONT hFont = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

    HFONT hBold = CreateFontW(-14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

    int yPos = 15;
    const int LBL_X = 15, LBL_W = 200, HK_X = 220, HK_W = 230, ROW_H = 28, GAP = 38;

    CreateHotkeyLabel(hwndSettings, L"Click a field, then press your shortcut.", LBL_X, yPos, DLG_W - 30, 20, hBold);
    yPos += 30;

    struct { const wchar_t* label; int id; HotkeyConfig* cfg; } rows[] = {
        { L"Clipboard Overlay:",        IDC_HK_OVERLAY,         &hotkeyConfig },
        { L"Snippets Overlay:",         IDC_HK_SNIPPETS,        &snippetsHotkey },
        { L"Copy Focused (UIA):",       IDC_HK_COPY_FOCUSED,    &copyFocusedHotkey },
        { L"Paste (Keystroke):",        IDC_HK_PASTE_FOCUSED,   &pasteFocusedHotkey },
        { L"Paste (Clipboard Swap):",   IDC_HK_PASTE_CLIPBOARD, &pasteClipboardHotkey },
    };

    for (auto& r : rows) {
        CreateHotkeyLabel(hwndSettings, r.label, LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
        HWND hHotkey = CreateHotkeyDisplay(hwndSettings, r.id, HotkeyToString(*r.cfg), HK_X, yPos, HK_W, ROW_H, hFont);
        SetWindowSubclass(hHotkey, SettingsHotkeyEditSubclass, 1, reinterpret_cast<DWORD_PTR>(this));
        yPos += GAP;
    }

    yPos += 14;
    CreateHotkeyLabel(hwndSettings, L"Overlay theme (AMOLED neon):", LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
    HWND hCombo = CreateWindowExW(0, L"COMBOBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
        HK_X, yPos, HK_W, 200, hwndSettings, (HMENU)(INT_PTR)IDC_THEME_COMBO,
        GetModuleHandle(nullptr), nullptr);
    if (hCombo) {
        SendMessageW(hCombo, WM_SETFONT, (WPARAM)hFont, TRUE);
        for (int t = 0; t < THEME_COUNT; t++)
            SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)kThemePresets[t].name);
        SendMessageW(hCombo, CB_SETCURSEL, (WPARAM)g_currentThemeId, 0);
    }
    yPos += GAP;

    // Font family: enumerate installed font face names (de-duplicated) and let the user
    // pick anything available on the system.
    CreateHotkeyLabel(hwndSettings, L"Overlay font:", LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
    HWND hFontCombo = CreateWindowExW(0, L"COMBOBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
        HK_X, yPos, HK_W, 300, hwndSettings, (HMENU)(INT_PTR)IDC_FONT_COMBO,
        GetModuleHandle(nullptr), nullptr);
    if (hFontCombo) {
        SendMessageW(hFontCombo, WM_SETFONT, (WPARAM)hFont, TRUE);
        struct EnumCtx { HWND combo; std::set<std::wstring> seen; };
        EnumCtx ctx{ hFontCombo, {} };
        HDC hdc = GetDC(nullptr);
        if (hdc) {
            LOGFONTW lf = {};
            lf.lfCharSet = DEFAULT_CHARSET;
            EnumFontFamiliesExW(hdc, &lf,
                [](const LOGFONTW* lpelfe, const TEXTMETRICW*, DWORD, LPARAM lp) -> int {
                    EnumCtx* c = (EnumCtx*)lp;
                    if (!lpelfe || lpelfe->lfFaceName[0] == L'@') return 1; // skip vertical fonts
                    std::wstring face = lpelfe->lfFaceName;
                    if (face.empty()) return 1;
                    if (c->seen.insert(face).second)
                        SendMessageW(c->combo, CB_ADDSTRING, 0, (LPARAM)face.c_str());
                    return 1;
                },
                (LPARAM)&ctx, 0);
            ReleaseDC(nullptr, hdc);
        }
        int sel = (int)SendMessageW(hFontCombo, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)g_themeFontFace.c_str());
        if (sel == CB_ERR)
            sel = (int)SendMessageW(hFontCombo, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)kDefaultThemeFontFace);
        if (sel != CB_ERR) SendMessageW(hFontCombo, CB_SETCURSEL, (WPARAM)sel, 0);
    }
    yPos += GAP;

    // Font color: small swatch + "Pick..." button that opens the standard color dialog.
    CreateHotkeyLabel(hwndSettings, L"Font color:", LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
    HWND hSwatch = CreateWindowExW(WS_EX_CLIENTEDGE, L"STATIC", L"",
        WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_OWNERDRAW,
        HK_X, yPos, 60, ROW_H, hwndSettings, (HMENU)(INT_PTR)IDC_FONT_COLOR_SWATCH,
        GetModuleHandle(nullptr), nullptr);
    (void)hSwatch;
    HWND hColorBtn = CreateWindowW(L"BUTTON", L"Pick color...",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        HK_X + 70, yPos, HK_W - 70, ROW_H, hwndSettings,
        (HMENU)(INT_PTR)IDC_FONT_COLOR_BTN, GetModuleHandle(nullptr), nullptr);
    if (hColorBtn) SendMessage(hColorBtn, WM_SETFONT, (WPARAM)hFont, TRUE);
    yPos += GAP;

    // History size: numeric edit (clamped on Save to [MIN_MAX_ITEMS, MAX_MAX_ITEMS]).
    CreateHotkeyLabel(hwndSettings, L"History size (items):", LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
    HWND hHistEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", std::to_wstring(maxItems).c_str(),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_LEFT,
        HK_X, yPos, HK_W, ROW_H, hwndSettings, (HMENU)(INT_PTR)IDC_HISTORY_SIZE,
        GetModuleHandle(nullptr), nullptr);
    if (hHistEdit) SendMessage(hHistEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    yPos += GAP;

    // Element colors: a row of clickable swatches (click = pick, double-click = reset that
    // element to the preset). Covers background, text, accent, selected text, border, dim.
    CreateHotkeyLabel(hwndSettings, L"Element colors:", LBL_X, yPos + 4, LBL_W, ROW_H, hFont);
    CreateHotkeyLabel(hwndSettings, L"(click = pick, double-click = reset)", LBL_X, yPos + 22, DLG_W - 30, 18, hFont);
    {
        const int SW_W = 64, SW_H = 26, SW_GAP = 6;
        int sx = HK_X;
        // Two swatches fit on the label line; the rest wrap to the line below.
        int swatchTop = yPos;
        for (int s = 0; s < TC_COUNT; s++) {
            int col = s % 3, rowIdx = s / 3;
            int px = HK_X + col * (SW_W + SW_GAP);
            int py = swatchTop + rowIdx * (SW_H + SW_GAP);
            CreateWindowExW(WS_EX_CLIENTEDGE, L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_NOTIFY | SS_OWNERDRAW,
                px, py, SW_W, SW_H, hwndSettings,
                (HMENU)(INT_PTR)(IDC_COLOR_SWATCH_FIRST + s),
                GetModuleHandle(nullptr), nullptr);
        }
        (void)sx;
    }
    yPos += GAP + 28;

    yPos += 6;
    const int BTN_W = 90, BTN_H = 30, BTN_GAP = 10;
    int btnX = DLG_W - 15 - BTN_W * 3 - BTN_GAP * 2 - 10;

    auto mkBtn = [&](int id, const wchar_t* text) {
        HWND b = CreateWindowW(L"BUTTON", text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            btnX, yPos, BTN_W, BTN_H, hwndSettings, (HMENU)(INT_PTR)id, GetModuleHandle(nullptr), nullptr);
        SendMessage(b, WM_SETFONT, (WPARAM)hFont, TRUE);
        btnX += BTN_W + BTN_GAP;
        return b;
    };

    mkBtn(IDC_BTN_DEFAULTS, L"Defaults");
    mkBtn(IDC_BTN_SAVE, L"Save");
    mkBtn(IDC_BTN_CANCEL, L"Cancel");

    ShowWindow(hwndSettings, SW_SHOW);
    UpdateWindow(hwndSettings);
}

LRESULT CALLBACK ClipboardManager::SettingsDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    if (!mgr) return DefWindowProc(hwnd, uMsg, wParam, lParam);

    switch (uMsg) {
    case WM_COMMAND: {
        int id = LOWORD(wParam);

        if (id == IDC_BTN_CANCEL) {
            DestroyWindow(hwnd);
            return 0;
        }

        if (id == IDC_BTN_DEFAULTS) {
            mgr->hotkeyConfig        = { MOD_CONTROL, VK_DECIMAL };
            mgr->snippetsHotkey      = { 0, 0 };
            mgr->copyFocusedHotkey   = { MOD_CONTROL, VK_F10 };
            mgr->pasteFocusedHotkey  = { MOD_CONTROL, VK_F11 };
            mgr->pasteClipboardHotkey = { MOD_CONTROL | MOD_SHIFT, VK_F11 };
            struct { int id; HotkeyConfig* cfg; } rows[] = {
                { IDC_HK_OVERLAY,         &mgr->hotkeyConfig },
                { IDC_HK_SNIPPETS,        &mgr->snippetsHotkey },
                { IDC_HK_COPY_FOCUSED,    &mgr->copyFocusedHotkey },
                { IDC_HK_PASTE_FOCUSED,   &mgr->pasteFocusedHotkey },
                { IDC_HK_PASTE_CLIPBOARD, &mgr->pasteClipboardHotkey },
            };
            for (auto& r : rows) {
                HWND hCtrl = GetDlgItem(hwnd, r.id);
                if (hCtrl) SetWindowTextW(hCtrl, HotkeyToString(*r.cfg).c_str());
            }
            HWND hCombo = GetDlgItem(hwnd, IDC_THEME_COMBO);
            if (hCombo) SendMessageW(hCombo, CB_SETCURSEL, (WPARAM)THEME_NEON_GREEN, 0);
            mgr->SetTheme(THEME_NEON_GREEN);
            // Reset font color override (back to preset) and font face (back to Consolas).
            mgr->SetThemeFontColor(kThemeFontColorPreset);
            mgr->SetThemeFontFace(kDefaultThemeFontFace);
            HWND hFontCombo = GetDlgItem(hwnd, IDC_FONT_COMBO);
            if (hFontCombo) {
                int sel = (int)SendMessageW(hFontCombo, CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)kDefaultThemeFontFace);
                if (sel != CB_ERR) SendMessageW(hFontCombo, CB_SETCURSEL, (WPARAM)sel, 0);
            }
            HWND hSwatch = GetDlgItem(hwnd, IDC_FONT_COLOR_SWATCH);
            if (hSwatch) InvalidateRect(hSwatch, nullptr, TRUE);
            // Reset all per-element color overrides back to the preset and refresh swatches.
            mgr->ResetThemeColorOverrides();
            for (int s = 0; s < TC_COUNT; s++) {
                HWND hsw = GetDlgItem(hwnd, IDC_COLOR_SWATCH_FIRST + s);
                if (hsw) InvalidateRect(hsw, nullptr, TRUE);
            }
            HWND hHistEdit = GetDlgItem(hwnd, IDC_HISTORY_SIZE);
            if (hHistEdit) SetWindowTextW(hHistEdit, std::to_wstring(DEFAULT_MAX_ITEMS).c_str());
            return 0;
        }

        // Live preview when the theme combo changes — easier than waiting for Save.
        if (id == IDC_THEME_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            HWND hCombo = (HWND)lParam;
            int sel = (int)SendMessageW(hCombo, CB_GETCURSEL, 0, 0);
            if (sel >= 0 && sel < THEME_COUNT)
                mgr->SetTheme(sel);
            // The preset's accent might have changed; if the swatch is in "preset" mode,
            // it needs a redraw to follow the new preset color.
            HWND hSwatch = GetDlgItem(hwnd, IDC_FONT_COLOR_SWATCH);
            if (hSwatch) InvalidateRect(hSwatch, nullptr, TRUE);
            return 0;
        }

        // Live preview when the font face combo changes.
        if (id == IDC_FONT_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            HWND hFontCombo = (HWND)lParam;
            int sel = (int)SendMessageW(hFontCombo, CB_GETCURSEL, 0, 0);
            if (sel >= 0) {
                int len = (int)SendMessageW(hFontCombo, CB_GETLBTEXTLEN, (WPARAM)sel, 0);
                if (len > 0 && len < LF_FACESIZE) {
                    wchar_t buf[LF_FACESIZE] = {};
                    SendMessageW(hFontCombo, CB_GETLBTEXT, (WPARAM)sel, (LPARAM)buf);
                    mgr->SetThemeFontFace(buf);
                }
            }
            return 0;
        }

        // Open the standard color dialog. Picking a color overrides the preset's TXT/SEL_BG.
        if (id == IDC_FONT_COLOR_BTN && HIWORD(wParam) == BN_CLICKED) {
            static COLORREF s_customColors[16] = {};
            CHOOSECOLORW cc = {};
            cc.lStructSize = sizeof(cc);
            cc.hwndOwner = hwnd;
            cc.lpCustColors = s_customColors;
            cc.rgbResult = (g_themeFontColor == kThemeFontColorPreset)
                ? kThemePresets[g_currentThemeId].txt
                : g_themeFontColor;
            cc.Flags = CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR;
            if (ChooseColorW(&cc)) {
                mgr->SetThemeFontColor(cc.rgbResult);
                HWND hSwatch = GetDlgItem(hwnd, IDC_FONT_COLOR_SWATCH);
                if (hSwatch) InvalidateRect(hSwatch, nullptr, TRUE);
            }
            return 0;
        }

        // Element color swatches: single-click opens the picker for that element;
        // double-click resets the element back to the active preset color.
        if (id >= IDC_COLOR_SWATCH_FIRST && id < IDC_COLOR_SWATCH_FIRST + TC_COUNT) {
            int slot = id - IDC_COLOR_SWATCH_FIRST;
            if (HIWORD(wParam) == STN_DBLCLK) {
                mgr->SetThemeColorOverride(slot, kColorUsePreset);
                for (int s = 0; s < TC_COUNT; s++) {
                    HWND hsw = GetDlgItem(hwnd, IDC_COLOR_SWATCH_FIRST + s);
                    if (hsw) InvalidateRect(hsw, nullptr, TRUE);
                }
                return 0;
            }
            if (HIWORD(wParam) == STN_CLICKED) {
                static COLORREF s_elemCustom[16] = {};
                CHOOSECOLORW cc = {};
                cc.lStructSize = sizeof(cc);
                cc.hwndOwner = hwnd;
                cc.lpCustColors = s_elemCustom;
                cc.rgbResult = EffectiveThemeColor(slot);
                cc.Flags = CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR;
                if (ChooseColorW(&cc)) {
                    mgr->SetThemeColorOverride(slot, cc.rgbResult);
                    // Other swatches (e.g. BG affects how every swatch frames) may shift.
                    for (int s = 0; s < TC_COUNT; s++) {
                        HWND hsw = GetDlgItem(hwnd, IDC_COLOR_SWATCH_FIRST + s);
                        if (hsw) InvalidateRect(hsw, nullptr, TRUE);
                    }
                }
                return 0;
            }
        }

        if (id == IDC_BTN_SAVE) {
            mgr->UnregisterHotkey();
            mgr->SaveHotkeyConfig();
            mgr->RegisterHotkey();
            HWND hCombo = GetDlgItem(hwnd, IDC_THEME_COMBO);
            if (hCombo) {
                int sel = (int)SendMessageW(hCombo, CB_GETCURSEL, 0, 0);
                if (sel >= 0 && sel < THEME_COUNT)
                    mgr->SetTheme(sel);
            }
            // History size: read, clamp, persist, and apply immediately.
            HWND hHist = GetDlgItem(hwnd, IDC_HISTORY_SIZE);
            if (hHist) {
                wchar_t buf[16] = {};
                GetWindowTextW(hHist, buf, 15);
                int v = _wtoi(buf);
                if (v < MIN_MAX_ITEMS) v = MIN_MAX_ITEMS;
                if (v > MAX_MAX_ITEMS) v = MAX_MAX_ITEMS;
                mgr->maxItems = v;
                SaveMaxItemsToRegistry(v);
                mgr->TrimHistory();
                mgr->FilterItems();
                mgr->UpdateListWindow();
            }
            // Font face/color have been applied + persisted live; nothing more to do for them.
            DestroyWindow(hwnd);
            return 0;
        }
        break;
    }

    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)lParam;
        if (dis && dis->CtlID == IDC_FONT_COLOR_SWATCH) {
            COLORREF c = (g_themeFontColor == kThemeFontColorPreset)
                ? kThemePresets[g_currentThemeId].txt
                : g_themeFontColor;
            // Black background frames the accent so it reads true to the AMOLED look.
            HBRUSH bg = CreateSolidBrush(RGB(0, 0, 0));
            FillRect(dis->hDC, &dis->rcItem, bg);
            DeleteObject(bg);
            RECT inner = dis->rcItem;
            InflateRect(&inner, -3, -3);
            HBRUSH fg = CreateSolidBrush(c);
            FillRect(dis->hDC, &inner, fg);
            DeleteObject(fg);
            return TRUE;
        }
        if (dis && dis->CtlID >= IDC_COLOR_SWATCH_FIRST && dis->CtlID < IDC_COLOR_SWATCH_FIRST + TC_COUNT) {
            int slot = dis->CtlID - IDC_COLOR_SWATCH_FIRST;
            COLORREF c = EffectiveThemeColor(slot);
            HBRUSH fg = CreateSolidBrush(c);
            FillRect(dis->hDC, &dis->rcItem, fg);
            DeleteObject(fg);
            // Overlay the element label in a contrasting color (luminance-based).
            int lum = (GetRValue(c) * 299 + GetGValue(c) * 587 + GetBValue(c) * 114) / 1000;
            SetTextColor(dis->hDC, lum > 140 ? RGB(0, 0, 0) : RGB(255, 255, 255));
            SetBkMode(dis->hDC, TRANSPARENT);
            std::wstring lbl = kColorLabels[slot];
            if (g_colorOverride[slot] != kColorUsePreset) lbl += L"*";  // marker: overridden
            DrawTextW(dis->hDC, lbl.c_str(), -1, &dis->rcItem, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            return TRUE;
        }
        break;
    }

    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        if (mgr->hwndSettings == hwnd)
            mgr->hwndSettings = nullptr;
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

static LONG WINAPI Clip2ExceptionHandler(LPEXCEPTION_POINTERS p) {
    wchar_t buf[256];
    swprintf_s(buf, L"clip2 crashed (exception 0x%08lX).", (unsigned long)p->ExceptionRecord->ExceptionCode);
    MessageBoxW(nullptr, buf, L"clip2 Error", MB_OK | MB_ICONERROR);
    return EXCEPTION_EXECUTE_HANDLER;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    SetUnhandledExceptionFilter(Clip2ExceptionHandler);
    try {
        ClipboardManager manager;
        
        if (!manager.Initialize()) {
            MessageBox(nullptr, L"Failed to initialize clip2", L"Error", MB_OK | MB_ICONERROR);
            return 1;
        }
        
        manager.Run();
    } catch (const std::exception& e) {
        std::string msg = std::string("clip2 crashed: ") + e.what();
        std::wstring wmsg(msg.begin(), msg.end());
        MessageBoxW(nullptr, wmsg.c_str(), L"clip2 Error", MB_OK | MB_ICONERROR);
        return 1;
    } catch (...) {
        MessageBox(nullptr, L"clip2 crashed with an unknown error during startup.", L"clip2 Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    return 0;
}

