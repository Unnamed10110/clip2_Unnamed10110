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
    
    ClipboardItem(UINT fmt, const std::vector<BYTE>& d) 
        : format(fmt), timestamp(std::chrono::system_clock::now()), thumbnail(nullptr), previewBitmap(nullptr), isImage(false), isVideo(false), formatName(L"Unknown Format"), fileType(L"Other"), preview(L"[Unknown]") {
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
    
private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK ListWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SearchEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SettingsDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
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
    void ShowListWindow();
    void HideListWindow();
    void UpdateListWindow();
    void ShowPreviewWindow(int itemIndex, int x, int y);
    void HidePreviewWindow();
    int GetItemAtPosition(int x, int y);
    void PasteItem(int index);
    void PasteMultipleItems();
    void ToggleMultiSelect(int filteredIndex);
    void ClearMultiSelection();
    void DeleteItem(int filteredIndex);
    void TransformTextItem(int filteredIndex, int transformType);
    bool IsTextItem(int actualIndex);
    void ProcessClipboard();
    void PlayClickSound();
    void ClearClipboardHistory();
    void FilterItems();
    void SetStartupWithWindows(bool enable);
    bool IsStartupWithWindows();
    void LoadHotkeyConfig();
    void SaveHotkeyConfig();
    void ShowSettingsDialog();
    
    struct HotkeyConfig {
        UINT modifiers;  // MOD_CONTROL, MOD_ALT, MOD_SHIFT, MOD_WIN
        UINT vkCode;     // Virtual key code
    };
    
    HWND hwndMain;
    HWND hwndList;
    HWND hwndPreview;
    HWND hwndSearch;
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
    std::set<int> multiSelectedIndices;  // Indices of items selected for multi-paste (filtered indices)
    bool isPasting;
    bool isProcessingClipboard;  // Prevent re-entrant clipboard processing
    HWND previousFocusWindow;
    int hoveredItemIndex;
    int selectedIndex;  // Currently selected item index (in filtered list)
    int multiSelectAnchor;  // Anchor index for Shift+Arrow multi-selection (-1 if no anchor)
    WNDPROC originalSearchEditProc;  // Original window procedure for search edit control
    HotkeyConfig hotkeyConfig;  // Current hotkey configuration
    HWND hwndSettings;  // Settings dialog window
    static const UINT WM_MOUSELEAVE_CUSTOM = WM_USER + 4;
    
    static const UINT WM_TRAYICON = WM_USER + 1;
    static const UINT WM_CLIPBOARD_HOTKEY = WM_USER + 2;
    static const UINT WM_PROCESS_CLIPBOARD = WM_USER + 3;
    static const int MAX_ITEMS = 100; // Reduced from 1000 to prevent memory issues
    static const int WINDOW_WIDTH = 600;
    static const int WINDOW_HEIGHT = 600;
    
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

