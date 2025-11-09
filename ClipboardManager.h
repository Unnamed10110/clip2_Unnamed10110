#pragma once

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <string>
#include <vector>
#include <memory>
#include <chrono>
#include <map>

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
        : format(fmt), timestamp(std::chrono::system_clock::now()), thumbnail(nullptr), previewBitmap(nullptr), isImage(false), isVideo(false) {
        formats[fmt] = d;
        formatName = GetFormatName(fmt);
        fileType = GetFileType(fmt);
        preview = GetPreview(d, fmt);
        GenerateThumbnail();
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
    void MonitorClipboard();
    void ProcessClipboard();
    void PlayClickSound();
    void ClearClipboardHistory();
    void FilterItems();
    void SetStartupWithWindows(bool enable);
    bool IsStartupWithWindows();
    
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
    bool isPasting;
    HWND previousFocusWindow;
    int hoveredItemIndex;
    int selectedIndex;  // Currently selected item index (in filtered list)
    WNDPROC originalSearchEditProc;  // Original window procedure for search edit control
    static const UINT WM_MOUSELEAVE_CUSTOM = WM_USER + 4;
    
    static const UINT WM_TRAYICON = WM_USER + 1;
    static const UINT WM_CLIPBOARD_HOTKEY = WM_USER + 2;
    static const UINT WM_PROCESS_CLIPBOARD = WM_USER + 3;
    static const int MAX_ITEMS = 1000;
    static const int WINDOW_WIDTH = 600;
    static const int WINDOW_HEIGHT = 600;
};

