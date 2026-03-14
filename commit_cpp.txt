#include "ClipboardManager.h"
#include <sstream>
#include <iomanip>
#include <ctime>
#include <fstream>
#include <algorithm>
#include <cwctype>
#include <cstring>
#include <shlobj.h>
#include <mmsystem.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <objbase.h>
#include <richedit.h>
#ifndef MSFTEDIT_CLASS
#define MSFTEDIT_CLASS L"RichEdit50W"
#endif
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")

// AS/400 5250 terminal theme (phosphor green on black)
namespace Theme5250 {
    const COLORREF BG        = RGB(0, 0, 0);
    const COLORREF TXT       = RGB(0, 255, 102);   // phosphor green
    const COLORREF SEL_BG    = RGB(0, 255, 102);   // reverse video
    const COLORREF SEL_FG    = RGB(0, 0, 0);       // text on selection
    const COLORREF BORDER    = RGB(0, 200, 80);
    const COLORREF DIM       = RGB(0, 128, 40);    // dividers, dimmer elements
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

// ClipboardManager implementation
ClipboardManager::ClipboardManager() 
    : hwndMain(nullptr), hwndList(nullptr), hwndPreview(nullptr), hwndSearch(nullptr), hwndSettings(nullptr), hwndSnippetsManager(nullptr), isRunning(false), listVisible(false), lastSequenceNumber(0), hKeyboardHook(nullptr), scrollOffset(0), itemsPerPage(10), numberInput(L""), searchText(L""), snippetsMode(false), lastSKeyTime(0), ignoreNextSChar(false), isPasting(false), isProcessingClipboard(false), lastPastedText(L""), previousFocusWindow(nullptr), hoveredItemIndex(-1), selectedIndex(0), multiSelectAnchor(-1), originalSearchEditProc(nullptr) {
    instance = this;
    ZeroMemory(&nid, sizeof(nid));
    // Default hotkey: Ctrl+NumPadDot
    hotkeyConfig.modifiers = MOD_CONTROL;
    hotkeyConfig.vkCode = VK_DECIMAL;
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
        static HFONT hSearchFont = CreateFontW(14, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, L"Consolas");
        SendMessage(hwndSearch, WM_SETFONT, (WPARAM)hSearchFont, TRUE);
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
            AppendMenu(hMenu, MF_STRING, 2, L"Exit");
            
            SetForegroundWindow(hwnd);
            
            // TrackPopupMenu needs special handling for proper message delivery
            DWORD cmd = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, 
                                       pt.x, pt.y, 0, hwnd, nullptr);
            
            // Post a fake message to ensure menu closes properly
            PostMessage(hwnd, WM_NULL, 0, 0);
            
            if (cmd == 1) {
                mgr->ShowListWindow();
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
        } else if (LOWORD(wParam) == 2) {
            mgr->Stop();
            PostQuitMessage(0);
        } else if (LOWORD(wParam) == 3) {
            // Toggle startup with Windows
            bool currentState = mgr->IsStartupWithWindows();
            mgr->SetStartupWithWindows(!currentState);
        }
        break;
        
    case WM_CLIPBOARD_HOTKEY:
        // Toggle list window visibility
        if (mgr->listVisible) {
            mgr->HideListWindow();
        } else {
            mgr->ShowListWindow();
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
                    // Check if foreground window is a child of our list
                    bool isChild = false;
                    if (fgWindow != nullptr) {
                        isChild = IsChild(mgr->hwndList, fgWindow) != FALSE;
                    }
                    if (!isChild) {
                        mgr->HideListWindow();
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
                    
                    // Draw background - 5250 black
                    FillRect(hdc, &rect, (HBRUSH)GetStockObject(BLACK_BRUSH));
                    
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
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        
        RECT rect;
        GetClientRect(hwnd, &rect);
        
        // Draw background - 5250 black
        FillRect(hdc, &rect, (HBRUSH)GetStockObject(BLACK_BRUSH));
        
        // Use monospace font for terminal look
        static HFONT hFont5250 = nullptr;
        if (!hFont5250) {
            hFont5250 = CreateFontW(14, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN,
                L"Consolas");
        }
        HGDIOBJ hOldFont = SelectObject(hdc, hFont5250);
        
        // Draw border - 5250 green (sharper corners for terminal look)
        HPEN hBorderPen = CreatePen(PS_SOLID, 2, Theme5250::BORDER);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hBorderPen);
        HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, 1, 1, rect.right - 1, rect.bottom - 1);
        SelectObject(hdc, hOldBrush);
        SelectObject(hdc, hOldPen);
        DeleteObject(hBorderPen);
        
        // Draw title - 5250 phosphor green
        SetTextColor(hdc, Theme5250::TXT);
        SetBkMode(hdc, TRANSPARENT);
        RECT titleRect = rect;
        titleRect.bottom = 30;
        if (mgr->snippetsMode) {
            std::wstring titleText = L"clip2 Snippets (Ctrl+Left = Clipboard | Ctrl+F search, Enter to paste)";
            DrawText(hdc, titleText.c_str(), -1, &titleRect, 
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        } else {
            RECT line1Rect = titleRect;
            line1Rect.bottom = 18;
            std::wstring line1 = L"clip2 Clipboard";
            if (!mgr->numberInput.empty()) line1 += L" [" + mgr->numberInput + L"]";
            DrawText(hdc, line1.c_str(), -1, &line1Rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            RECT hintRect = titleRect;
            hintRect.top = 16;
            hintRect.bottom = 30;
            DrawText(hdc, L"Ctrl+Right = Snippets | Ctrl+Left = Clipboard", -1, &hintRect, 
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
        
        // Draw search box border (if visible) - 5250 green
        if (mgr->hwndSearch && IsWindowVisible(mgr->hwndSearch)) {
            RECT searchRect = {5, 30, WINDOW_WIDTH - 5, 55};
            HPEN hPen = CreatePen(PS_SOLID, 1, Theme5250::BORDER);
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
            Rectangle(hdc, searchRect.left, searchRect.top, searchRect.right, searchRect.bottom);
            SelectObject(hdc, hOldBrush);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);
        }
        
        // Calculate scroll info
        int itemHeight = 50;
        int totalItems = mgr->snippetsMode ? (int)mgr->filteredSnippetIndices.size() : (int)mgr->filteredIndices.size();
        int visibleHeight = rect.bottom - rect.top - 60; // Subtract title height (30) and search box height (30)
        mgr->itemsPerPage = visibleHeight / itemHeight;
        int maxScroll = std::max(0, totalItems - mgr->itemsPerPage);
        
        // Clamp scroll offset
        if (mgr->scrollOffset > maxScroll) {
            mgr->scrollOffset = maxScroll;
        }
        if (mgr->scrollOffset < 0) {
            mgr->scrollOffset = 0;
        }
        
        // Draw items (only visible ones)
        int yPos = 60;
        int startIndex = mgr->scrollOffset;
        int endIndex = std::min(startIndex + mgr->itemsPerPage, totalItems);
        
        if (mgr->snippetsMode) {
            for (int i = startIndex; i < endIndex && yPos + itemHeight <= rect.bottom - 10; i++) {
                if (i >= (int)mgr->filteredSnippetIndices.size()) break;
                int actualIndex = mgr->filteredSnippetIndices[i];
                const auto& snip = mgr->snippets[actualIndex];
                
                RECT itemRect = {5, yPos, rect.right - 5, yPos + itemHeight - 2};
                bool isSelected = (i == mgr->selectedIndex && mgr->selectedIndex >= 0);
                
                if (isSelected) {
                    HBRUSH selectedBrush = CreateSolidBrush(Theme5250::SEL_BG);
                    FillRect(hdc, &itemRect, selectedBrush);
                    DeleteObject(selectedBrush);
                } else {
                    HBRUSH itemBrush = CreateSolidBrush(Theme5250::BG);
                    FillRect(hdc, &itemRect, itemBrush);
                    DeleteObject(itemBrush);
                }
                
                int leftOffset = 5;
                std::wstring num = std::to_wstring(i + 1) + L".";
                RECT numRect = itemRect;
                numRect.left += leftOffset;
                numRect.right = numRect.left + 35;
                SetTextColor(hdc, isSelected ? Theme5250::SEL_FG : Theme5250::TXT);
                SetBkMode(hdc, TRANSPARENT);
                DrawText(hdc, num.c_str(), -1, &numRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                
                RECT previewRect = itemRect;
                previewRect.left += leftOffset + 40;
                previewRect.right -= 10;
                std::wstring preview = snip.name;
                std::wstring dispContent = snip.contentPlain.empty() ? snip.content : snip.contentPlain;
                if (!dispContent.empty()) {
                    std::wstring contentPreview = dispContent.length() > 50 ? dispContent.substr(0, 47) + L"..." : dispContent;
                    preview += L" - " + contentPreview;
                }
                SetTextColor(hdc, isSelected ? Theme5250::SEL_FG : Theme5250::TXT);
                DrawText(hdc, preview.c_str(), -1, &previewRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                
                if (i < endIndex - 1) {
                    HPEN hPen = CreatePen(PS_SOLID, 1, Theme5250::DIM);
                    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                    MoveToEx(hdc, itemRect.left, itemRect.bottom - 1, nullptr);
                    LineTo(hdc, itemRect.right, itemRect.bottom - 1);
                    SelectObject(hdc, hOldPen);
                    DeleteObject(hPen);
                }
                yPos += itemHeight;
            }
        } else {
        for (int i = startIndex; i < endIndex && yPos + itemHeight <= rect.bottom - 10; i++) {
            if (i >= (int)mgr->filteredIndices.size()) break;
            int actualIndex = mgr->filteredIndices[i];
            const auto& item = mgr->clipboardHistory[actualIndex];
            
            RECT itemRect = {5, yPos, rect.right - 5, yPos + itemHeight - 2};
            bool isMultiSelected = mgr->multiSelectedIndices.find(i) != mgr->multiSelectedIndices.end();
            bool isSelected = (i == mgr->selectedIndex && mgr->selectedIndex >= 0);
            
            if (isMultiSelected || isSelected) {
                HBRUSH selectedBrush = CreateSolidBrush(Theme5250::SEL_BG);
                FillRect(hdc, &itemRect, selectedBrush);
                DeleteObject(selectedBrush);
            } else {
                HBRUSH itemBrush = CreateSolidBrush(Theme5250::BG);
                FillRect(hdc, &itemRect, itemBrush);
                DeleteObject(itemBrush);
            }
            
            int leftOffset = 5;
            if (item->thumbnail) {
                HDC hdcMem = CreateCompatibleDC(hdc);
                SelectObject(hdcMem, item->thumbnail);
                int thumbSize = 40;
                int thumbY = yPos + (itemHeight - thumbSize) / 2;
                BitBlt(hdc, leftOffset + 5, thumbY, thumbSize, thumbSize, hdcMem, 0, 0, SRCCOPY);
                DeleteDC(hdcMem);
                leftOffset += thumbSize + 10;
            } else if (item->isImage || item->isVideo) {
                // Draw placeholder icon for images/videos without thumbnail - 5250 green
                RECT iconRect = {leftOffset + 5, yPos + 5, leftOffset + 45, yPos + 45};
                HBRUSH iconBrush = CreateSolidBrush(Theme5250::DIM);
                FillRect(hdc, &iconRect, iconBrush);
                DeleteObject(iconBrush);
                leftOffset += 50;
            }
            
            // Draw number - 5250 reverse video when selected
            std::wstring num = std::to_wstring(i + 1) + L".";
            RECT numRect = itemRect;
            numRect.left += leftOffset;
            numRect.right = numRect.left + 35;
            SetTextColor(hdc, (isMultiSelected || isSelected) ? Theme5250::SEL_FG : Theme5250::TXT);
            SetBkMode(hdc, TRANSPARENT);
            DrawText(hdc, num.c_str(), -1, &numRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            
            // Draw preview - 5250 theme
            RECT previewRect = itemRect;
            previewRect.left += leftOffset + 40;
            previewRect.right -= 180;
            SetTextColor(hdc, (isMultiSelected || isSelected) ? Theme5250::SEL_FG : Theme5250::TXT);
            DrawText(hdc, item->preview.c_str(), -1, &previewRect, 
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            
            // Draw format info - White text on black, black text on white
            std::wstring info = item->formatName + L" - " + item->fileType;
            RECT infoRect = itemRect;
            infoRect.left = infoRect.right - 175;
            SetTextColor(hdc, (isMultiSelected || isSelected) ? Theme5250::SEL_FG : Theme5250::TXT);
            
            // Format date
            auto time_t = std::chrono::system_clock::to_time_t(item->timestamp);
            std::wstringstream ss;
            ss << std::put_time(std::localtime(&time_t), L"%H:%M:%S");
            info += L" [" + ss.str() + L"]";
            
            DrawText(hdc, info.c_str(), -1, &infoRect, 
                    DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            
            if (i < endIndex - 1) {
                HPEN hPen = CreatePen(PS_SOLID, 1, Theme5250::DIM);
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, itemRect.left, itemRect.bottom - 1, nullptr);
                LineTo(hdc, itemRect.right, itemRect.bottom - 1);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
            }
            
            yPos += itemHeight;
        }
        }
        
        SelectObject(hdc, hOldFont);
        EndPaint(hwnd, &ps);
        return 0;
    }
    
    case WM_KEYDOWN: {
        // Ctrl+Right = Snippets mode, Ctrl+Left = Clipboard mode
        bool ctrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrlPressed && (wParam == VK_RIGHT || wParam == VK_LEFT)) {
            bool wantSnippets = (wParam == VK_RIGHT);
            if (mgr->snippetsMode != wantSnippets) {
                mgr->snippetsMode = wantSnippets;
                mgr->numberInput.clear();
                mgr->scrollOffset = 0;
                if (mgr->snippetsMode) {
                    mgr->searchText.clear();
                    if (mgr->hwndSearch) {
                        SetWindowText(mgr->hwndSearch, L"");
                        SendMessage(mgr->hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search snippets (Ctrl+F) | Arrow keys to move, Enter to paste");
                    }
                    mgr->FilterSnippets();
                    mgr->selectedIndex = (mgr->filteredSnippetIndices.empty() ? -1 : 0);
                    SetFocus(mgr->hwndList);
                } else {
                    mgr->FilterItems();
                    if (mgr->hwndSearch) {
                        SendMessage(mgr->hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search... (Ctrl+F)");
                    }
                }
                mgr->UpdateListWindow();
            }
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
            
            if (mgr->snippetsMode) {
                mgr->HidePreviewWindow();
            } else if (itemIndex >= 0 && itemIndex < (int)mgr->filteredIndices.size()) {
                int actualIndex = mgr->filteredIndices[itemIndex];
                const auto& item = mgr->clipboardHistory[actualIndex];
                if (item->isImage || item->isVideo) {
                    if (mgr->hoveredItemIndex != itemIndex) {
                        mgr->hoveredItemIndex = itemIndex;
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
        mgr->hoveredItemIndex = -1;
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
            // Style the search edit control - 5250 theme
            HDC hdcEdit = (HDC)wParam;
            SetTextColor(hdcEdit, Theme5250::TXT);
            SetBkColor(hdcEdit, Theme5250::BG);
            static HBRUSH hDarkBrush = CreateSolidBrush(Theme5250::BG);
            return (LRESULT)hDarkBrush;
        }
    
    case WM_KILLFOCUS:
        {
            // Hide list when it loses focus (unless focus is moving to a child control)
            HWND hwndGettingFocus = (HWND)wParam;
            
            // Check if focus is moving to a child of the list window (like search box)
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            
            if (!isChildWindow && hwndGettingFocus != hwnd) {
                // Focus is moving to another window (not a child) - hide the list
                mgr->HideListWindow();
            } else {
                // Focus is moving to a child control or staying in list - keep window on top
                SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            }
        }
        break;
    
    case WM_ACTIVATE:
        if (wParam == WA_INACTIVE) {
            // Window is becoming inactive - hide the list immediately
            // Check if focus is going to a child control first
            HWND hwndGettingFocus = GetFocus();
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            
            if (!isChildWindow) {
                mgr->HideListWindow();
            }
        } else {
            SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        }
        break;
    
    case WM_NCACTIVATE:
        // Hide list when window loses non-client activation (clicking title bar or outside)
        if (wParam == FALSE) {
            // Check if focus is going to a child control
            HWND hwndGettingFocus = GetFocus();
            bool isChildWindow = false;
            if (hwndGettingFocus != nullptr) {
                isChildWindow = IsChild(hwnd, hwndGettingFocus) != FALSE;
            }
            
            if (!isChildWindow) {
                mgr->HideListWindow();
            }
        }
        // Still call DefWindowProc to allow normal activation behavior
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
    // Legacy function - not used, kept for compatibility
}

void ClipboardManager::UnregisterHotkey() {
    // Legacy function - not used, kept for compatibility
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
        
        // Check for configured hotkey modifiers
        bool hasCtrl = (mgr->hotkeyConfig.modifiers & MOD_CONTROL) != 0;
        bool hasAlt = (mgr->hotkeyConfig.modifiers & MOD_ALT) != 0;
        bool hasShift = (mgr->hotkeyConfig.modifiers & MOD_SHIFT) != 0;
        bool hasWin = (mgr->hotkeyConfig.modifiers & MOD_WIN) != 0;
        
        bool isCtrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        bool isAltPressed = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
        bool isShiftPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
        bool isWinPressed = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;
        
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
    
    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

void ClipboardManager::ShowListWindow() {
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
    // Start in clipboard mode (not snippets)
    snippetsMode = false;
    // Initialize filtered indices (show all items initially)
    FilterItems();
    // Reset selection to first item (if any items exist)
    if (!filteredIndices.empty()) {
        selectedIndex = 0;
    } else {
        selectedIndex = -1;
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
    int clickedY = y - 60; // Subtract title height (30) and search box height (30)
    
    if (clickedY < 0) return -1;
    
    int filteredIndex = scrollOffset + (clickedY / itemHeight);
    
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
    // index is a filtered index, convert to actual index
    if (index < 0 || index >= (int)filteredIndices.size()) {
        return;
    }
    
    int actualIndex = filteredIndices[index];
    if (actualIndex < 0 || actualIndex >= (int)clipboardHistory.size()) {
        return;
    }
    
    const auto& item = clipboardHistory[actualIndex];
    
    // Check if Ctrl is held - if so, paste as plain text only
    bool pasteAsPlainText = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
    
    // Set flag to ignore clipboard changes we're about to make
    isPasting = true;
    
    // Small delay to ensure clipboard is ready
    Sleep(10);
    
    if (OpenClipboard(hwndMain)) {
        EmptyClipboard();
        
        bool success = false;
        
        if (pasteAsPlainText) {
            // Paste as plain text only - extract plain text from the item
            std::wstring plainText;
            
            // Try to extract plain text - check formats in order of preference
            const std::vector<BYTE>* textData = nullptr;
            
            // Try Unicode text first
            textData = item->GetFormatData(CF_UNICODETEXT);
            if (textData && textData->size() >= sizeof(wchar_t)) {
                size_t len = (textData->size() / sizeof(wchar_t)) - 1; // Exclude null terminator
                plainText = std::wstring((wchar_t*)textData->data(), len);
            } else {
                // Try ANSI text
                textData = item->GetFormatData(CF_TEXT);
                if (textData && textData->size() > 0) {
                    size_t len = textData->size() - 1; // Exclude null terminator
                    std::string ansiText((char*)textData->data(), len);
                    int wideLen = MultiByteToWideChar(CP_ACP, 0, ansiText.c_str(), -1, nullptr, 0);
                    if (wideLen > 0) {
                        plainText.resize(wideLen - 1);
                        MultiByteToWideChar(CP_ACP, 0, ansiText.c_str(), -1, &plainText[0], wideLen);
                    }
                } else {
                    // For other formats (RTF, HTML, etc.), try to extract text from preview
                    // Use the preview string which should contain plain text representation
                    plainText = item->preview;
                }
            }
            
            if (!plainText.empty()) {
                // Set as Unicode text
                size_t dataSize = (plainText.length() + 1) * sizeof(wchar_t);
                HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize);
                if (hMem) {
                    wchar_t* pMem = (wchar_t*)GlobalLock(hMem);
                    if (pMem) {
                        wcscpy_s(pMem, plainText.length() + 1, plainText.c_str());
                        GlobalUnlock(hMem);
                        
                        if (SetClipboardData(CF_UNICODETEXT, hMem)) {
                            success = true;
                        } else {
                            GlobalFree(hMem);
                        }
                    } else {
                        GlobalFree(hMem);
                    }
                }
            }
        } else {
            // Paste with ALL original formats to preserve formatting (bold, colors, sizes, etc.)
            // This ensures rich text, RTF, HTML, and all other formats are restored
            bool firstFormat = true;
            for (const auto& formatPair : item->formats) {
                UINT fmt = formatPair.first;
                const std::vector<BYTE>& formatData = formatPair.second;
                
                HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, formatData.size());
                if (hMem) {
                    void* pMem = GlobalLock(hMem);
                    if (pMem) {
                        memcpy(pMem, formatData.data(), formatData.size());
                        GlobalUnlock(hMem);
                        
                        // SetClipboardData takes ownership of hMem on success, so don't free it
                        if (SetClipboardData(fmt, hMem)) {
                            if (firstFormat) {
                                success = true;
                                firstFormat = false;
                            }
                            // Note: hMem is now owned by the clipboard, don't free it
                        } else {
                            // Only free if SetClipboardData failed
                            GlobalFree(hMem);
                        }
                    } else {
                        GlobalFree(hMem);
                    }
                }
            }
        }
        
        CloseClipboard();
        
        if (success) {
            // Store the text we're about to paste so we can ignore it if it's copied back
            // Extract text from the item for comparison
            if (!pasteAsPlainText) {
                // For full format paste, extract text from CF_UNICODETEXT if available
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
                            lastPastedText = std::wstring(textPtr, actualLen);
                        }
                    }
                } else {
                    // Fallback to preview text
                    lastPastedText = item->preview;
                }
            }
            
            // Update sequence number AFTER setting clipboard - this is the sequence number of our paste
            lastSequenceNumber = GetClipboardSequenceNumber();
            
            // Restore focus to previous window before pasting
            if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
                SetForegroundWindow(previousFocusWindow);
                SetFocus(previousFocusWindow);
                Sleep(50); // Give window time to receive focus
            }
            
            // Small delay before pasting
            Sleep(50);
            
            // Simulate paste (Ctrl+V)
            keybd_event(VK_CONTROL, 0, 0, 0);
            keybd_event('V', 0, 0, 0);
            keybd_event('V', 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            
            // Keep isPasting flag true longer to prevent re-adding the pasted item
            // Some applications copy the pasted content back to clipboard, so we need to wait
            Sleep(500); // Increased delay to ignore clipboard updates from paste operation
            
            isPasting = false;
        } else {
            isPasting = false;
        }
    } else {
        isPasting = false;
    }
}

void ClipboardManager::PasteMultipleItems() {
    if (multiSelectedIndices.empty()) {
        return;
    }
    
    // Sort indices to paste in order (top to bottom)
    std::vector<int> sortedIndices(multiSelectedIndices.begin(), multiSelectedIndices.end());
    std::sort(sortedIndices.begin(), sortedIndices.end());
    
    // Combine all text items into a single string with newlines
    std::wstring combinedText;
    bool hasTextItems = false;
    
    for (size_t i = 0; i < sortedIndices.size(); i++) {
        int filteredIndex = sortedIndices[i];
        int actualIndex = filteredIndices[filteredIndex];
        if (actualIndex >= 0 && actualIndex < (int)clipboardHistory.size()) {
            const auto& item = clipboardHistory[actualIndex];
            
            // Extract text from the item
            std::wstring itemText;
            const std::vector<BYTE>* textData = item->GetFormatData(CF_UNICODETEXT);
            if (textData && textData->size() >= sizeof(wchar_t)) {
                size_t len = textData->size() / sizeof(wchar_t);
                if (len > 0) {
                    const wchar_t* textPtr = (const wchar_t*)textData->data();
                    // Find actual length (up to null terminator)
                    size_t actualLen = len;
                    for (size_t j = 0; j < len; j++) {
                        if (textPtr[j] == L'\0') {
                            actualLen = j;
                            break;
                        }
                    }
                    if (actualLen > 0) {
                        itemText = std::wstring(textPtr, actualLen);
                        hasTextItems = true;
                    }
                }
            } else {
                // Try ANSI text
                textData = item->GetFormatData(CF_TEXT);
                if (textData && textData->size() > 0) {
                    const char* ansiText = (const char*)textData->data();
                    size_t len = textData->size();
                    size_t actualLen = len;
                    for (size_t j = 0; j < len; j++) {
                        if (ansiText[j] == '\0') {
                            actualLen = j;
                            break;
                        }
                    }
                    if (actualLen > 0) {
                        std::string ansiStr(ansiText, actualLen);
                        int wideLen = MultiByteToWideChar(CP_ACP, 0, ansiStr.c_str(), (int)ansiStr.length(), nullptr, 0);
                        if (wideLen > 0) {
                            itemText.resize(wideLen);
                            MultiByteToWideChar(CP_ACP, 0, ansiStr.c_str(), (int)ansiStr.length(), &itemText[0], wideLen);
                            hasTextItems = true;
                        }
                    }
                }
            }
            
            if (!itemText.empty()) {
                combinedText += itemText;
                // Add newline after each item except the last
                if (i < sortedIndices.size() - 1) {
                    combinedText += L"\n";
                }
            }
        }
    }
    
    // If we have combined text, paste it as a single operation
    if (hasTextItems && !combinedText.empty()) {
        // Set flag to ignore clipboard changes we're about to make
        isPasting = true;
        
        // Restore focus to previous window before pasting
        if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
            SetForegroundWindow(previousFocusWindow);
            SetFocus(previousFocusWindow);
            Sleep(150); // Give window time to receive focus
        }
        
        // Ensure clipboard is available (retry logic)
        bool clipboardReady = false;
        int retries = 0;
        while (!clipboardReady && retries < 10) {
            if (OpenClipboard(hwndMain)) {
                clipboardReady = true;
            } else {
                Sleep(10);
                retries++;
            }
        }
        
        if (clipboardReady) {
            EmptyClipboard();
            
            // Set as Unicode text
            size_t dataSize = (combinedText.length() + 1) * sizeof(wchar_t);
            HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize);
            if (hMem) {
                wchar_t* pMem = (wchar_t*)GlobalLock(hMem);
                if (pMem) {
                    wcscpy_s(pMem, combinedText.length() + 1, combinedText.c_str());
                    GlobalUnlock(hMem);
                    
                        if (SetClipboardData(CF_UNICODETEXT, hMem)) {
                        // Store the text we're about to paste so we can ignore it if it's copied back
                        lastPastedText = combinedText;
                        
                        // Update sequence number AFTER setting clipboard
                        lastSequenceNumber = GetClipboardSequenceNumber();
                        
                        CloseClipboard();
                        
                        // Small delay before pasting
                        Sleep(150);
                        
                        // Simulate paste (Ctrl+V) - single paste for all items
                        keybd_event(VK_CONTROL, 0, 0, 0);
                        keybd_event('V', 0, 0, 0);
                        keybd_event('V', 0, KEYEVENTF_KEYUP, 0);
                        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
                        
                        // Keep isPasting flag true longer to prevent re-adding the pasted item
                        // Some applications copy the pasted content back to clipboard, so we need to wait
                        Sleep(500); // Increased delay to ignore clipboard updates from paste operation
                        
                        isPasting = false;
                    } else {
                        GlobalFree(hMem);
                        CloseClipboard();
                    }
                } else {
                    GlobalFree(hMem);
                    CloseClipboard();
                }
            } else {
                CloseClipboard();
                isPasting = false;
            }
        } else {
            isPasting = false;
        }
    }
    
    // Clear multi-selection after pasting
    ClearMultiSelection();
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

void ClipboardManager::ProcessClipboard() {
    // Prevent re-entrant calls
    if (isProcessingClipboard) {
        return;
    }
    
    isProcessingClipboard = true;
    
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
                    const int MAX_FORMATS_PER_ITEM = 5; // Reduced to 5 formats max
                    const size_t MAX_FORMAT_SIZE_PER_ITEM = 10 * 1024 * 1024; // 10MB per item total (reduced)
                    
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
                            
                            // Limit to MAX_ITEMS
                            while (clipboardHistory.size() > (size_t)MAX_ITEMS) {
                                clipboardHistory.pop_back();
                            }
                            
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

static std::wstring GetClickSoundPath() {
    static std::wstring path;
    static bool initialized = false;
    if (!initialized) {
        initialized = true;
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(nullptr, exePath, MAX_PATH);
        wchar_t* lastSlash = wcsrchr(exePath, L'\\');
        if (lastSlash) *(lastSlash + 1) = L'\0';
        path = std::wstring(exePath) + L"click.mp3";
    }
    return path;
}

void ClipboardManager::WarmUpClickSound() {
    std::wstring soundPath = GetClickSoundPath();
    std::wstring openCmd = L"open \"" + soundPath + L"\" type MPEGVideo alias clicksnd_warmup";
    if (mciSendString(openCmd.c_str(), nullptr, 0, nullptr) == 0) {
        mciSendString(L"close clicksnd_warmup", nullptr, 0, nullptr);
    }
}

void ClipboardManager::PlayClickSound() {
    static bool aliasOpen = false;
    std::wstring soundPath = GetClickSoundPath();
    
    if (aliasOpen) {
        mciSendString(L"stop clicksnd", nullptr, 0, nullptr);
        mciSendString(L"close clicksnd", nullptr, 0, nullptr);
        aliasOpen = false;
    }
    
    std::wstring openCmd = L"open \"" + soundPath + L"\" type MPEGVideo alias clicksnd";
    if (mciSendString(openCmd.c_str(), nullptr, 0, nullptr) == 0) {
        aliasOpen = true;
        mciSendString(L"play clicksnd", nullptr, 0, nullptr);
    }
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
    // Clear all items - destructors will clean up bitmaps and memory
    clipboardHistory.clear();
    // Force cleanup of related state
    filteredIndices.clear();
    scrollOffset = 0;
    hoveredItemIndex = -1;
    selectedIndex = 0;
    numberInput.clear();
    searchText.clear();
    
    // Force garbage collection by clearing vectors
    std::vector<std::unique_ptr<ClipboardItem>>().swap(clipboardHistory);
    std::vector<int>().swap(filteredIndices);
    filteredIndices.clear();
    if (hwndSearch) {
        SetWindowText(hwndSearch, L"");
    }
    UpdateListWindow();
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
        char magic[4];
        DWORD read;
        if (!ReadFile(h, magic, 4, &read, nullptr) || read != 4 || memcmp(magic, "CLP2", 4) != 0) { CloseHandle(h); return; }
        DWORD version, count;
        if (!ReadFile(h, &version, 4, &read, nullptr) || read != 4 || version != 1) { CloseHandle(h); return; }
        if (!ReadFile(h, &count, 4, &read, nullptr) || read != 4) { CloseHandle(h); return; }
        count = std::min(count, (DWORD)300);
        for (DWORD i = 0; i < count; i++) {
            DWORD fmt, size;
            if (!ReadFile(h, &fmt, 4, &read, nullptr) || read != 4) break;
            if (!ReadFile(h, &size, 4, &read, nullptr) || read != 4) break;
            if (size == 0 || size > 500000) continue;
            if (fmt != CF_UNICODETEXT && fmt != CF_TEXT) continue;
            std::vector<BYTE> data(size);
            if (!ReadFile(h, data.data(), size, &read, nullptr) || read != size) break;
            if (fmt == CF_UNICODETEXT && size >= sizeof(wchar_t)) {
                data.push_back(0);
                data.push_back(0);
            } else if (fmt == CF_TEXT && size >= 1) {
                data.push_back(0);
            }
            try {
                auto item = std::make_unique<ClipboardItem>(fmt, data);
                clipboardHistory.push_back(std::move(item));
            } catch (...) {}
        }
        CloseHandle(h);
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
        HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;
        DWORD written;
        WriteFile(h, "CLP2", 4, &written, nullptr);
        DWORD version = 1;
        WriteFile(h, &version, 4, &written, nullptr);
        size_t toSave = std::min(clipboardHistory.size(), (size_t)300);
        DWORD count = (DWORD)toSave;
        WriteFile(h, &count, 4, &written, nullptr);
        for (size_t i = 0; i < toSave; i++) {
            const auto& item = clipboardHistory[i];
            const std::vector<BYTE>* data = item->GetFormatData(CF_UNICODETEXT);
            UINT fmt = CF_UNICODETEXT;
            if (!data || data->empty()) {
                data = item->GetFormatData(CF_TEXT);
                fmt = CF_TEXT;
            }
            if (!data || data->empty()) continue;
            size_t size = data->size();
            if (fmt == CF_UNICODETEXT && size >= sizeof(wchar_t)) {
                size_t len = size / sizeof(wchar_t);
                const wchar_t* p = (const wchar_t*)data->data();
                while (len > 0 && p[len - 1] == 0) len--;
                size = len * sizeof(wchar_t);
            } else if (fmt == CF_TEXT && size >= 1) {
                size_t len = 0;
                const char* p = (const char*)data->data();
                while (len < size && p[len]) len++;
                size = len;
            }
            if (size == 0 || size > 500000) continue;
            DWORD ufmt = (DWORD)fmt, usize = (DWORD)size;
            WriteFile(h, &ufmt, 4, &written, nullptr);
            WriteFile(h, &usize, 4, &written, nullptr);
            WriteFile(h, data->data(), (DWORD)size, &written, nullptr);
        }
        CloseHandle(h);
    } catch (...) {}
}

void ClipboardManager::FilterItems() {
    filteredIndices.clear();
    ClearMultiSelection(); // Clear multi-selection when filtering
    
    if (searchText.empty()) {
        // No filter - show all items
        for (size_t i = 0; i < clipboardHistory.size(); i++) {
            filteredIndices.push_back((int)i);
        }
    } else {
        // Filter items based on search text (case-insensitive)
        std::wstring searchLower = searchText;
        std::transform(searchLower.begin(), searchLower.end(), searchLower.begin(), ::towlower);
        
        for (size_t i = 0; i < clipboardHistory.size(); i++) {
            const auto& item = clipboardHistory[i];
            
            // Search in full text content (handles line breaks, not just first 50 chars)
            std::wstring fullText = item->GetFullSearchableText();
            if (!fullText.empty()) {
                std::wstring fullLower = fullText;
                std::transform(fullLower.begin(), fullLower.end(), fullLower.begin(), ::towlower);
                if (fullLower.find(searchLower) != std::wstring::npos) {
                    filteredIndices.push_back((int)i);
                    continue;
                }
            }
            
            // Fallback: search in preview, format name, and file type (for images, files, etc.)
            std::wstring previewLower = item->preview;
            std::transform(previewLower.begin(), previewLower.end(), previewLower.begin(), ::towlower);
            std::wstring formatLower = item->formatName;
            std::transform(formatLower.begin(), formatLower.end(), formatLower.begin(), ::towlower);
            std::wstring fileTypeLower = item->fileType;
            std::transform(fileTypeLower.begin(), fileTypeLower.end(), fileTypeLower.begin(), ::towlower);
            
            if (previewLower.find(searchLower) != std::wstring::npos ||
                formatLower.find(searchLower) != std::wstring::npos ||
                fileTypeLower.find(searchLower) != std::wstring::npos) {
                filteredIndices.push_back((int)i);
            }
        }
    }
    
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
        
        for (size_t i = 0; i < snippets.size(); i++) {
            std::wstring nameLower = snippets[i].name;
            std::transform(nameLower.begin(), nameLower.end(), nameLower.begin(), ::towlower);
            std::wstring contentLower = snippets[i].content;
            std::transform(contentLower.begin(), contentLower.end(), contentLower.begin(), ::towlower);
            
            if (nameLower.find(searchLower) != std::wstring::npos ||
                contentLower.find(searchLower) != std::wstring::npos) {
                filteredSnippetIndices.push_back((int)i);
            }
        }
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

void ClipboardManager::LoadHotkeyConfig() {
    HKEY hKey;
    LONG result = RegOpenKeyExW(
        HKEY_CURRENT_USER,
        L"Software\\clip2",
        0,
        KEY_READ,
        &hKey
    );
    
    if (result == ERROR_SUCCESS) {
        DWORD modifiers = 0;
        DWORD vkCode = 0;
        DWORD dataSize = sizeof(DWORD);
        DWORD type = REG_DWORD;
        
        // Read modifiers
        result = RegQueryValueExW(
            hKey,
            L"HotkeyModifiers",
            nullptr,
            &type,
            (LPBYTE)&modifiers,
            &dataSize
        );
        
        if (result == ERROR_SUCCESS && type == REG_DWORD) {
            hotkeyConfig.modifiers = modifiers;
        }
        
        // Read vkCode
        dataSize = sizeof(DWORD);
        result = RegQueryValueExW(
            hKey,
            L"HotkeyVkCode",
            nullptr,
            &type,
            (LPBYTE)&vkCode,
            &dataSize
        );
        
        if (result == ERROR_SUCCESS && type == REG_DWORD) {
            hotkeyConfig.vkCode = vkCode;
        }
        
        RegCloseKey(hKey);
    }
}

void ClipboardManager::SaveHotkeyConfig() {
    HKEY hKey;
    LONG result = RegCreateKeyExW(
        HKEY_CURRENT_USER,
        L"Software\\clip2",
        0,
        nullptr,
        REG_OPTION_NON_VOLATILE,
        KEY_WRITE,
        nullptr,
        &hKey,
        nullptr
    );
    
    if (result == ERROR_SUCCESS) {
        DWORD modifiers = hotkeyConfig.modifiers;
        DWORD vkCode = hotkeyConfig.vkCode;
        
        RegSetValueExW(
            hKey,
            L"HotkeyModifiers",
            0,
            REG_DWORD,
            (BYTE*)&modifiers,
            sizeof(DWORD)
        );
        
        RegSetValueExW(
            hKey,
            L"HotkeyVkCode",
            0,
            REG_DWORD,
            (BYTE*)&vkCode,
            sizeof(DWORD)
        );
        
        RegCloseKey(hKey);
    }
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
    std::wstring text = ExpandSnippetPlaceholders(snip.content);
    std::wstring plainText = snip.contentPlain.empty() ? text : ExpandSnippetPlaceholders(snip.contentPlain);
    if (text.empty() && plainText.empty()) return;

    isPasting = true;
    Sleep(10);

    UINT cfRtf = RegisterClipboardFormatA("Rich Text Format");
    bool isRtf = IsRtfContent(text);

    if (OpenClipboard(hwndMain)) {
        EmptyClipboard();
        bool success = false;
        if (isRtf && cfRtf != 0) {
            std::string rtfBytes;
            for (wchar_t wc : text) rtfBytes += (char)(wc & 0xFF);
            HGLOBAL hRtf = GlobalAlloc(GMEM_MOVEABLE, rtfBytes.size() + 1);
            if (hRtf) {
                char* pRtf = (char*)GlobalLock(hRtf);
                if (pRtf) {
                    memcpy(pRtf, rtfBytes.c_str(), rtfBytes.size() + 1);
                    GlobalUnlock(hRtf);
                    if (SetClipboardData(cfRtf, hRtf)) success = true;
                    else GlobalFree(hRtf);
                } else GlobalFree(hRtf);
            }
        }
        if (!plainText.empty()) {
            size_t dataSize = (plainText.length() + 1) * sizeof(wchar_t);
            HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize);
            if (hMem) {
                wchar_t* pMem = (wchar_t*)GlobalLock(hMem);
                if (pMem) {
                    wcscpy_s(pMem, plainText.length() + 1, plainText.c_str());
                    GlobalUnlock(hMem);
                    if (SetClipboardData(CF_UNICODETEXT, hMem)) success = true;
                } else GlobalFree(hMem);
            } else GlobalFree(hMem);
        } else if (!success && !text.empty()) {
            size_t dataSize = (text.length() + 1) * sizeof(wchar_t);
            HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize);
            if (hMem) {
                wchar_t* pMem = (wchar_t*)GlobalLock(hMem);
                if (pMem) {
                    wcscpy_s(pMem, text.length() + 1, text.c_str());
                    GlobalUnlock(hMem);
                    if (SetClipboardData(CF_UNICODETEXT, hMem)) success = true;
                } else GlobalFree(hMem);
            } else GlobalFree(hMem);
        }
        CloseClipboard();
        if (success || !plainText.empty() || !text.empty()) {
            lastPastedText = plainText.empty() ? text : plainText;
            lastSequenceNumber = GetClipboardSequenceNumber();
            if (previousFocusWindow && previousFocusWindow != hwndList && IsWindow(previousFocusWindow)) {
                SetForegroundWindow(previousFocusWindow);
                SetFocus(previousFocusWindow);
                Sleep(50);
            }
            Sleep(50);
            keybd_event(VK_CONTROL, 0, 0, 0);
            keybd_event('V', 0, 0, 0);
            keybd_event('V', 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            Sleep(500);
        }
    }
    isPasting = false;
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

void ClipboardManager::ShowSettingsDialog() {
    // Show current hotkey and allow basic configuration
    std::wstring hotkeyText = L"Current hotkey: ";
    if (hotkeyConfig.modifiers & MOD_CONTROL) hotkeyText += L"Ctrl+";
    if (hotkeyConfig.modifiers & MOD_ALT) hotkeyText += L"Alt+";
    if (hotkeyConfig.modifiers & MOD_SHIFT) hotkeyText += L"Shift+";
    if (hotkeyConfig.modifiers & MOD_WIN) hotkeyText += L"Win+";
    
    // Get key name
    wchar_t keyName[256] = {0};
    UINT scanCode = MapVirtualKeyW(hotkeyConfig.vkCode, MAPVK_VK_TO_VSC);
    if (scanCode != 0) {
        LONG lParam = (scanCode << 16);
        // Add extended key flag if needed (for numpad keys)
        if (hotkeyConfig.vkCode == VK_DECIMAL || hotkeyConfig.vkCode == VK_NUMPAD0 || 
            (hotkeyConfig.vkCode >= VK_NUMPAD1 && hotkeyConfig.vkCode <= VK_NUMPAD9)) {
            lParam |= (1 << 24); // Extended key flag
        }
        if (GetKeyNameTextW(lParam, keyName, 256)) {
            hotkeyText += keyName;
        } else {
            hotkeyText += L"Key " + std::to_wstring(hotkeyConfig.vkCode);
        }
    } else {
        hotkeyText += L"Key " + std::to_wstring(hotkeyConfig.vkCode);
    }
    
    hotkeyText += L"\n\nPress a key combination now to set it as the new hotkey.\n";
    hotkeyText += L"Click OK, then press your desired key combination within 3 seconds.\n";
    hotkeyText += L"Supported modifiers: Ctrl, Alt, Shift, Win";
    
    int result = MessageBox(hwndMain, hotkeyText.c_str(), L"clip2 Settings - Hotkey Configuration", MB_OKCANCEL | MB_ICONINFORMATION);
    
    if (result == IDOK) {
        // Wait for user to press a key combination
        MessageBox(hwndMain, L"Please press your desired hotkey combination now...", L"clip2 Settings", MB_OK | MB_ICONINFORMATION);
        
        // Simple hotkey capture - wait for key press
        // Note: This is a simplified implementation
        // In a full implementation, you'd use a hook or dialog with proper key capture
        
        // For now, show message that user can configure via registry
        std::wstring configText = L"Hotkey configuration via dialog is coming soon.\n\n";
        configText += L"For now, you can configure the hotkey by editing the registry:\n\n";
        configText += L"Location: HKEY_CURRENT_USER\\Software\\clip2\n\n";
        configText += L"Values:\n";
        configText += L"  HotkeyModifiers (DWORD):\n";
        configText += L"    MOD_CONTROL = 2\n";
        configText += L"    MOD_ALT = 1\n";
        configText += L"    MOD_SHIFT = 4\n";
        configText += L"    MOD_WIN = 8\n";
        configText += L"    Combine with + (e.g., MOD_CONTROL | MOD_SHIFT = 6)\n\n";
        configText += L"  HotkeyVkCode (DWORD): Virtual key code\n";
        configText += L"    Examples:\n";
        configText += L"      VK_DECIMAL (NumPad .) = 110\n";
        configText += L"      VK_F1 = 112\n";
        configText += L"      'V' = 0x56\n";
        configText += L"      Space = 0x20\n\n";
        configText += L"After editing, restart clip2 for changes to take effect.";
        
        MessageBox(hwndMain, configText.c_str(), L"clip2 Settings", MB_OK | MB_ICONINFORMATION);
    }
}

LRESULT CALLBACK ClipboardManager::SettingsDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ClipboardManager* mgr = instance;
    
    switch (uMsg) {
    case WM_CLOSE:
        if (mgr && mgr->hwndSettings == hwnd) {
            mgr->hwndSettings = nullptr;
        }
        DestroyWindow(hwnd);
        return 0;
        
    case WM_DESTROY:
        if (mgr && mgr->hwndSettings == hwnd) {
            mgr->hwndSettings = nullptr;
        }
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

