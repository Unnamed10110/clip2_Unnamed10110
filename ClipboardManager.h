#pragma once

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <string>
#include <vector>
#include <memory>
#include <chrono>
#include <map>
#include <set>

struct ClipboardItem {
    UINT format;  // Primary format (for display/compatibility)
    std::map<UINT, std::vector<BYTE>> formats;  // All formats stored
    std::wstring preview;
    std::wstring formatName;
    std::wstring fileType;
    std::chrono::system_clock::time_point timestamp;
    HBITMAP thumbnail;
    HBITMAP previewBitmap;  // Larger preview for hover
    bool isImage;
    bool isVideo;
    bool pinned;  // Pinned/favorite items stay on top, survive Clear, and always persist
    
    ClipboardItem(UINT fmt, const std::vector<BYTE>& d) 
        : format(fmt), timestamp(std::chrono::system_clock::now()), thumbnail(nullptr), previewBitmap(nullptr), isImage(false), isVideo(false), pinned(false), formatName(L"Unknown Format"), fileType(L"Other"), preview(L"[Unknown]") {
        // CRITICAL: Store data first, before any processing
        if (d.empty() || d.size() > 200 * 1024 * 1024) {
            return;
        }
        
        // Store format data immediately - this must succeed
        formats[fmt] = d;
        
        // Get format name - simple, safe operation
        if (fmt == CF_TEXT) formatName = L"Text";
        else if (fmt == CF_UNICODETEXT) formatName = L"Unicode Text";
        else if (fmt == CF_OEMTEXT) formatName = L"OEM Text";
        else if (fmt == CF_BITMAP) formatName = L"Bitmap";
        else if (fmt == CF_DIB) formatName = L"DIB";
        else if (fmt == CF_DIBV5) formatName = L"DIB v5";
        else if (fmt == CF_HDROP) formatName = L"File Drop";
        else {
            // Try to get format name, but don't crash if it fails
            try {
                wchar_t name[256] = {0};
                if (GetClipboardFormatNameW(fmt, name, 256)) {
                    formatName = name;
                }
            } catch (...) {
                formatName = L"Unknown Format";
            }
        }
        
        // Get file type - simple checks only
        if (fmt == CF_TEXT || fmt == CF_UNICODETEXT || fmt == CF_OEMTEXT) {
            fileType = L"Text";
        } else if (fmt == CF_BITMAP || fmt == CF_DIB || fmt == CF_DIBV5) {
            fileType = L"Image";
            isImage = true;
        } else if (fmt == CF_HDROP) {
            fileType = L"Files";
        } else {
            fileType = L"Other";
        }
        
        // Get preview - ONLY for text formats, skip complex processing
        if (fmt == CF_UNICODETEXT && d.size() >= sizeof(wchar_t) && d.size() < 100000) {
            try {
                size_t len = d.size() / sizeof(wchar_t);
                if (len > 0 && len < 10000) {
                    const wchar_t* text = (const wchar_t*)d.data();
                    if (text) {
                        size_t previewLen = std::min(len, (size_t)50);
                        preview.assign(text, previewLen);
                        if (len > 50) preview += L"...";
                    }
                }
            } catch (...) {
                preview = L"[Unicode Text]";
            }
        } else if (fmt == CF_TEXT && d.size() > 0 && d.size() < 100000) {
            try {
                size_t len = std::min(d.size(), (size_t)50);
                const char* text = (const char*)d.data();
                if (text) {
                    std::string str(text, len);
                    preview = std::wstring(str.begin(), str.end());
                    if (d.size() > 50) preview += L"...";
                }
            } catch (...) {
                preview = L"[Text]";
            }
        } else {
            preview = L"[" + formatName + L"]";
        }
        
        // Skip thumbnail generation for text - it's not needed and can cause crashes
        if (fmt != CF_TEXT && fmt != CF_UNICODETEXT && fmt != CF_OEMTEXT) {
            try {
                GenerateThumbnail();
            } catch (...) {
                // Ignore thumbnail errors
            }
        }
    }
    
    // Add additional format
    void AddFormat(UINT fmt, const std::vector<BYTE>& d) {
        formats[fmt] = d;
    }
    
    // Get data for a specific format
    const std::vector<BYTE>* GetFormatData(UINT fmt) const {
        auto it = formats.find(fmt);
        if (it != formats.end()) {
            return &it->second;
        }
        return nullptr;
    }
    
    // Get full text content for search (handles line breaks; empty for non-text)
    std::wstring GetFullSearchableText() const;
    
    ~ClipboardItem() {
        if (thumbnail) {
            DeleteObject(thumbnail);
            thumbnail = nullptr;
        }
        if (previewBitmap) {
            DeleteObject(previewBitmap);
            previewBitmap = nullptr;
        }
    }
    
    HBITMAP GetPreviewBitmap();  // Generate larger preview on demand
    
private:
    std::wstring GetFormatName(UINT fmt);
    std::wstring GetFileType(UINT fmt);
    std::wstring GetPreview(const std::vector<BYTE>& data, UINT fmt);
    void GenerateThumbnail();
    HBITMAP CreateBitmapFromData();
};

class ClipboardManager {
public:
    ClipboardManager();
    ~ClipboardManager();
    
    bool Initialize();
    void Run();
    void Stop();

    struct HotkeyConfig {
        UINT modifiers;  // MOD_CONTROL, MOD_ALT, MOD_SHIFT, MOD_WIN
        UINT vkCode;     // Virtual key code
    };

    static std::wstring HotkeyToString(const HotkeyConfig& hk);
    static bool HotkeyFromKeyMessage(UINT vk, HotkeyConfig& out);
    
private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK ListWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SearchEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SettingsDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SettingsHotkeyEditSubclass(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
        UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    static ClipboardManager* instance;
    
    void CreateTrayIcon();
    void UpdateTrayIcon();
    void RemoveTrayIcon();
    void RegisterHotkey();
    void UnregisterHotkey();
    void InstallKeyboardHook();
    void UninstallKeyboardHook();
    static LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
    void FocusListWindow();
    void ShowListWindow(bool startInSnippetsMode = false);
    void SetListSnippetsMode(bool wantSnippets);
    void HideListWindow();
    void UpdateListWindow(bool includeInactive = true);
    void SaveOverlayPosition();                  // Persist main-list top-left + keep pinned panel aligned
    void ClampOverlayPosition(int& mainX, int& mainY) const; // Keep restored position on-screen
    // Left "pinned" panel support. The live member state (filteredIndices, selectedIndex,
    // scrollOffset, searchText, snippetsMode, multi-select...) always belongs to whichever
    // pane is focused; the other pane keeps a lightweight snapshot for rendering.
    HWND ActiveListHwnd();                       // hwndPinned when activeIsPinned else hwndList
    bool IsOverlayWindow(HWND h);                // any overlay window / child / preview
    void SwitchActivePane(bool toPinned, bool focusListWindow = true);  // move state/focus between panels
    void EnsureActivePane(bool wantPinned, bool focusListWindow = true); // switch only if needed
    void RefreshInactivePane();                  // recompute the non-focused pane's snapshot list
    void ComputeFilteredForPane(bool pinnedOnly, const std::wstring& search, std::vector<int>& out);
    void ShowPreviewWindow(int itemIndex, int x, int y);
    void HidePreviewWindow();
    int GetItemAtPosition(int x, int y);
    void PasteItem(int index);
    void PasteMultipleItems();
    void ToggleMultiSelect(int filteredIndex);
    void ClearMultiSelection();
    void DeleteItem(int filteredIndex);
    void TransformTextItem(int filteredIndex, int transformType);
    void TogglePin(int filteredIndex);            // Pin/unpin an item (filtered index)
    void TrimHistory();                           // Enforce maxItems cap without evicting pinned items
    void CopyItemAsPlainText(int filteredIndex);  // Quick action: strip formatting -> plain Unicode on clipboard
    void OpenItemUrl(int filteredIndex);          // Quick action: ShellExecute the item's URL
    void SaveImageItemToFile(int filteredIndex);  // Quick action: write image item to a .bmp file
    bool IsTextItem(int actualIndex);
    void ProcessClipboard();
    void PlayClickSound();
    void WarmUpClickSound();
    void ClearClipboardHistory();
    void LoadClipboardHistory();
    void SaveClipboardHistory();
    void FilterItems();
    void FilterSnippets();
    void SetStartupWithWindows(bool enable);
    bool IsStartupWithWindows();
    void LoadHotkeyConfig();
    void SaveHotkeyConfig();
    void ShowSettingsDialog();
    void SetTheme(int themeId);                     // Apply + persist a theme preset and repaint UI.
    void SetThemeFontColor(COLORREF color);          // Override (or clear, via sentinel) overlay font color.
    void SetThemeFontFace(const std::wstring& face); // Switch overlay/search font face and repaint.
    void SetThemeColorOverride(int slot, COLORREF color); // Override (or clear) one overlay element color.
    void ResetThemeColorOverrides();                 // Clear all per-element color overrides.
    void RefreshThemeVisuals();                      // Re-apply class brushes + repaint after a color change.
    // Read text from focused control via UI Automation (when app does not put anything on clipboard)
    bool CopyFromFocusedControlViaUIA();
    // Universal fallback: synthetic Ctrl+C (then Ctrl+A + Ctrl+C) against the focused app; captured text lands on the clipboard.
    bool CopyFocusedViaSyntheticCopy(std::wstring& outText);
    // Shared tail for captured text: add to history, optionally set the system clipboard, play the click sound.
    void CommitCapturedText(const std::wstring& text, bool setClipboard = true);
    // Paste from history into focused control: useClipboardSwap=false sends plain text as Unicode keystrokes (Ctrl+F11).
    // useClipboardSwap=true puts item formats/text on clipboard and simulates Ctrl+V (Ctrl+Shift+F11).
    bool PasteToFocusedControlWithoutClipboard(bool useClipboardSwap);
    
    HWND hwndMain;
    HWND hwndList;
    HWND hwndPinned;          // Left panel showing only pinned items (same class/proc as hwndList)
    HWND hwndPreview;
    HWND hwndSearch;          // Active pane's search edit (points at hwndMainSearch or hwndPinnedSearch)
    HWND hwndMainSearch;      // Search edit child of hwndList
    HWND hwndPinnedSearch;    // Search edit child of hwndPinned
    bool activeIsPinned;      // Which pane currently owns the live state / keyboard focus
    struct PaneState {
        std::vector<int> filteredIndices;
        int selectedIndex = 0;
        int scrollOffset = 0;
        std::wstring searchText;
    };
    PaneState inactivePane;   // Snapshot of the non-focused pane (for rendering + restore on switch)
    DWORD overlayShownTick;   // Tick when overlay was last shown (grace period before focus-hide)
    bool overlayGotForeground;// True once the overlay actually became foreground after showing
    bool hasSavedOverlayPos;  // True when overlayPosX/Y come from a prior drag or registry
    int overlayPosX;          // Last main-list top-left X (virtual-screen coords)
    int overlayPosY;          // Last main-list top-left Y
    NOTIFYICONDATA nid;
    UINT wmTaskbarCreated;
    bool isRunning;
    bool listVisible;
    std::vector<std::unique_ptr<ClipboardItem>> clipboardHistory;
    UINT lastSequenceNumber;
    HHOOK hKeyboardHook;
    int scrollOffset;
    int itemsPerPage;
    std::wstring numberInput;
    std::wstring searchText;
    std::vector<int> filteredIndices;  // Indices of items matching search
    std::vector<int> filteredSnippetIndices;  // Indices of snippets matching search (when snippetsMode)
    std::set<int> multiSelectedIndices;  // Indices of items selected for multi-paste (filtered indices)
    bool snippetsMode;  // When true, overlay shows snippets instead of clipboard
    DWORD lastSKeyTime;  // For detecting "ss" double-press
    bool ignoreNextSChar;  // Ignore next 's' to prevent stray char when entering snippets via "ss"
    bool isPasting;
    bool isProcessingClipboard;  // Prevent re-entrant clipboard processing
    std::wstring lastPastedText;  // Store last pasted text to ignore it if it's copied back
    HWND previousFocusWindow;
    int hoveredItemIndex;
    int selectedIndex;  // Currently selected item index (in filtered list)
    int multiSelectAnchor;  // Anchor index for Shift+Arrow multi-selection (-1 if no anchor)
    WNDPROC originalSearchEditProc;  // Original window procedure for search edit control
    HotkeyConfig hotkeyConfig;           // Overlay toggle
    HotkeyConfig snippetsHotkey;         // Snippets overlay toggle
    HotkeyConfig copyFocusedHotkey;      // Copy from focused control (UIA)
    HotkeyConfig pasteFocusedHotkey;     // Keystroke injection paste
    HotkeyConfig pasteClipboardHotkey;   // Clipboard swap + Ctrl+V paste
    DWORD lastHotkeyTick;      // Debounce: last time overlay hotkey was handled
    HWND hwndSettings;  // Settings dialog window
    std::map<UINT, std::vector<BYTE>> immediateClipboardSnapshot;  // Captured in WM_CLIPBOARDUPDATE before app can clear
    bool hasImmediateClipboardSnapshot;
    struct Snippet {
        std::wstring name;
        std::wstring content;       // RTF or plain text
        std::wstring contentPlain;  // Plain text fallback (for RTF snippets)
    };
    
    void LoadSnippets();
    void SaveSnippets();
    void PasteSnippet(int index);
    std::wstring ExpandSnippetPlaceholders(const std::wstring& content);
    void ShowSnippetsManagerDialog();
    
    // Bypass copy blocks: capture clipboard immediately when it changes so we have content before apps clear it
    bool TryCaptureClipboardImmediately();
    void ProcessClipboardFromSnapshot();
    
    static LRESULT CALLBACK SnippetsManagerProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    
    std::vector<Snippet> snippets;
    HWND hwndSnippetsManager;
    
    static const UINT WM_MOUSELEAVE_CUSTOM = WM_USER + 4;
    
    static const UINT WM_TRAYICON = WM_USER + 1;
    static const UINT WM_CLIPBOARD_HOTKEY = WM_USER + 2;
    static const UINT WM_PROCESS_CLIPBOARD = WM_USER + 3;
    static const UINT WM_SNIPPETS_OVERLAY_HOTKEY = WM_USER + 5;
    static const UINT WM_DISMISS_OVERLAY = WM_USER + 6;  // e.g. Esc from low-level hook
    static const int HOTKEY_ID_OVERLAY = 1;
    static const int HOTKEY_ID_COPY_FOCUSED = 2;  // Ctrl+F10: copy from focused control (UIA)
    static const int HOTKEY_ID_PASTE_FOCUSED = 3;           // Ctrl+F11: keystroke injection (bypass Trillex / paste blocks)
    static const int HOTKEY_ID_PASTE_FOCUSED_CLIPBOARD = 4;  // Ctrl+Shift+F11: clipboard swap + Ctrl+V (VS Code, etc.)
    static const int DEFAULT_MAX_ITEMS = 300;   // Default history cap when no user override exists
    static const int MIN_MAX_ITEMS = 10;        // Lower clamp for the configurable history size
    static const int MAX_MAX_ITEMS = 2000;      // Upper clamp for the configurable history size
    int maxItems;                               // Runtime-configurable history cap (registry: MaxItems)
    static const int WINDOW_WIDTH = 600;
    static const int WINDOW_HEIGHT = 600;
    static const int PINNED_WIDTH = 320;   // Width of the left pinned panel
    static const int PANEL_GAP = 12;       // Gap between the pinned panel and the main list
    
    // Text transformation types
    enum TextTransform {
        TRANSFORM_UPPERCASE = 200,
        TRANSFORM_LOWERCASE = 201,
        TRANSFORM_TITLE_CASE = 202,
        TRANSFORM_REMOVE_LINE_BREAKS = 203,
        TRANSFORM_TRIM_WHITESPACE = 204,
        TRANSFORM_PLAIN_TEXT = 205  // Remove formatting, keep only plain text
    };
};

