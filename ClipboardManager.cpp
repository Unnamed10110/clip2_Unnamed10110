#include "ClipboardManager.h"
#include <sstream>
#include <iomanip>
#include <fstream>
#include <algorithm>
#include <shlobj.h>
#include <mmsystem.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <objbase.h>
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")

ClipboardManager* ClipboardManager::instance = nullptr;

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
    // DISABLED: Preview bitmap generation causes memory issues
    // Just return thumbnail to prevent crashes
    return thumbnail;
}

// ClipboardManager implementation
ClipboardManager::ClipboardManager() 
    : hwndMain(nullptr), hwndList(nullptr), hwndPreview(nullptr), hwndSearch(nullptr), isRunning(false), listVisible(false), lastSequenceNumber(0), hKeyboardHook(nullptr), scrollOffset(0), itemsPerPage(10), numberInput(L""), searchText(L""), isPasting(false), isProcessingClipboard(false), previousFocusWindow(nullptr), hoveredItemIndex(-1), selectedIndex(0), originalSearchEditProc(nullptr) {
    instance = this;
    ZeroMemory(&nid, sizeof(nid));
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
    wcList.hbrBackground = CreateSolidBrush(RGB(0, 0, 0)); // Black OLED background
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
    wcPreview.hbrBackground = CreateSolidBrush(RGB(0, 0, 0)); // Black OLED background
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
    
    // Set rounded window region for rounded borders
    HRGN hRgn = CreateRoundRectRgn(0, 0, WINDOW_WIDTH + 1, WINDOW_HEIGHT + 1, 15, 15);
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
        // Set placeholder text hint
        SendMessage(hwndSearch, EM_SETCUEBANNER, TRUE, (LPARAM)L"Search... (Ctrl+F)");
        // Subclass the edit control to handle Enter key
        originalSearchEditProc = (WNDPROC)SetWindowLongPtr(hwndSearch, GWLP_WNDPROC, (LONG_PTR)SearchEditProc);
        ShowWindow(hwndSearch, SW_HIDE);  // Hide initially, show when list is visible
    }
    
    // Create preview window (initially hidden) with rounded borders
    hwndPreview = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED,
        L"ClipboardManagerPreviewClass",
        L"Preview",
        WS_POPUP, // Removed WS_BORDER - we'll draw rounded border ourselves
        CW_USEDEFAULT, CW_USEDEFAULT,
        300, 300,
        nullptr, nullptr, GetModuleHandle(nullptr), nullptr
    );
    
    if (hwndPreview) {
        // Set rounded window region for preview window
        HRGN hPreviewRgn = CreateRoundRectRgn(0, 0, 301, 301, 15, 15);
        SetWindowRgn(hwndPreview, hPreviewRgn, TRUE);
        DeleteObject(hPreviewRgn);
        
        ShowWindow(hwndPreview, SW_HIDE);
        SetLayeredWindowAttributes(hwndPreview, 0, 255, LWA_ALPHA);
    }
    
    CreateTrayIcon();
    InstallKeyboardHook();
    
    wmTaskbarCreated = RegisterWindowMessage(L"TaskbarCreated");
    lastSequenceNumber = GetClipboardSequenceNumber();
    
    // Register for clipboard update notifications (instant notification when clipboard changes)
    AddClipboardFormatListener(hwndMain);
    
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
            POINT pt;
            GetCursorPos(&pt);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, 1, L"Show Clipboard");
            AppendMenu(hMenu, MF_SEPARATOR, 0, nullptr);
            
            // Add "Start with Windows" option with checkmark
            bool startupEnabled = mgr->IsStartupWithWindows();
            AppendMenu(hMenu, startupEnabled ? (MF_STRING | MF_CHECKED) : MF_STRING, 3, L"Start with Windows");
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
            // Play sound immediately when clipboard changes
            mgr->PlayClickSound();
            
            // Update sequence number
            mgr->lastSequenceNumber = GetClipboardSequenceNumber();
            
            // Post message to process clipboard asynchronously
            PostMessage(hwnd, WM_PROCESS_CLIPBOARD, 0, 0);
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
                    
                    // Draw background - Black OLED theme
                    FillRect(hdc, &rect, (HBRUSH)GetStockObject(BLACK_BRUSH));
                    
                    // Draw rounded border - Bright red
                    HPEN hPreviewBorderPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));
                    HPEN hOldPreviewPen = (HPEN)SelectObject(hdc, hPreviewBorderPen);
                    HBRUSH hOldPreviewBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
                    RoundRect(hdc, 1, 1, rect.right - 1, rect.bottom - 1, 15, 15);
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
                        
                        // Scale to fit preview window (max 350x350)
                        int maxSize = 350;
                        int srcW = bm.bmWidth;
                        int srcH = bm.bmHeight;
                        float scale = std::min((float)maxSize / srcW, (float)maxSize / srcH);
                        int dstW = (int)(srcW * scale);
                        int dstH = (int)(srcH * scale);
                        int offsetX = (rect.right - dstW) / 2;
                        int offsetY = (rect.bottom - dstH) / 2;
                        
                        // Use better scaling
                        SetStretchBltMode(hdc, HALFTONE);
                        StretchBlt(hdc, offsetX, offsetY, dstW, dstH, hdcMem, 0, 0, srcW, srcH, SRCCOPY);
                        
                        DeleteDC(hdcMem);
                    } else if (item->isImage || item->isVideo) {
                        // Draw placeholder - Darker colors for OLED theme
                        RECT centerRect = rect;
                        centerRect.left += 50;
                        centerRect.top += 50;
                        centerRect.right -= 50;
                        centerRect.bottom -= 50;
                        HBRUSH brush = CreateSolidBrush(item->isVideo ? RGB(80, 40, 40) : RGB(40, 60, 80));
                        FillRect(hdc, &centerRect, brush);
                        DeleteObject(brush);
                        
                        SetTextColor(hdc, RGB(200, 200, 200));
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
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        
        RECT rect;
        GetClientRect(hwnd, &rect);
        
        // Draw background - Black OLED theme
        FillRect(hdc, &rect, (HBRUSH)GetStockObject(BLACK_BRUSH));
        
        // Draw rounded border - Bright red
        HPEN hBorderPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));
        HPEN hOldPen = (HPEN)SelectObject(hdc, hBorderPen);
        HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        RoundRect(hdc, 1, 1, rect.right - 1, rect.bottom - 1, 15, 15);
        SelectObject(hdc, hOldBrush);
        SelectObject(hdc, hOldPen);
        DeleteObject(hBorderPen);
        
        // Draw title - Light gray text on black
        SetTextColor(hdc, RGB(220, 220, 220));
        SetBkMode(hdc, TRANSPARENT);
        RECT titleRect = rect;
        titleRect.bottom = 30;
        std::wstring titleText = L"clip2 (Ctrl+NumPadDot to hide, Type number + Enter to paste)";
        if (!mgr->numberInput.empty()) {
            titleText += L" [" + mgr->numberInput + L"]";
        }
        DrawText(hdc, titleText.c_str(), -1, &titleRect, 
                DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        
        // Draw search box border (if visible) - Bright red
        if (mgr->hwndSearch && IsWindowVisible(mgr->hwndSearch)) {
            RECT searchRect = {5, 30, WINDOW_WIDTH - 5, 55};
            HPEN hPen = CreatePen(PS_SOLID, 1, RGB(255, 0, 0));
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
            RoundRect(hdc, searchRect.left, searchRect.top, searchRect.right, searchRect.bottom, 5, 5);
            SelectObject(hdc, hOldBrush);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);
        }
        
        // Calculate scroll info (no visible scrollbar, but we still track scrolling)
        int itemHeight = 50;
        int totalItems = (int)mgr->filteredIndices.size();
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
        int yPos = 60; // Start below title (30) and search box (30)
        int startIndex = mgr->scrollOffset;
        int endIndex = std::min(startIndex + mgr->itemsPerPage, totalItems);
        
        for (int i = startIndex; i < endIndex && yPos + itemHeight <= rect.bottom - 10; i++) {
            if (i >= (int)mgr->filteredIndices.size()) break;
            int actualIndex = mgr->filteredIndices[i];
            const auto& item = mgr->clipboardHistory[actualIndex];
            
            // Draw item background - Dark gray for items, bright red for selected
            RECT itemRect = {5, yPos, rect.right - 5, yPos + itemHeight - 2};
            if (i == mgr->selectedIndex && mgr->selectedIndex >= 0) {
                // Highlight selected item - Bright red
                HBRUSH selectedBrush = CreateSolidBrush(RGB(255, 0, 0));
                FillRect(hdc, &itemRect, selectedBrush);
                DeleteObject(selectedBrush);
            } else {
                // Regular item - Dark gray
                HBRUSH itemBrush = CreateSolidBrush(RGB(20, 20, 20));
                FillRect(hdc, &itemRect, itemBrush);
                DeleteObject(itemBrush);
            }
            
            // Draw thumbnail if available (for images/videos)
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
                // Draw placeholder icon for images/videos without thumbnail
                RECT iconRect = {leftOffset + 5, yPos + 5, leftOffset + 45, yPos + 45};
                HBRUSH iconBrush = CreateSolidBrush(item->isVideo ? RGB(200, 100, 100) : RGB(100, 150, 200));
                FillRect(hdc, &iconRect, iconBrush);
                DeleteObject(iconBrush);
                leftOffset += 50;
            }
            
            // Draw number with highlight (use filtered position, not actual index) - Bright red
            std::wstring num = std::to_wstring(i + 1) + L".";
            RECT numRect = itemRect;
            numRect.left += leftOffset;
            numRect.right = numRect.left + 35;
            SetTextColor(hdc, RGB(255, 0, 0));
            SetBkMode(hdc, TRANSPARENT);
            DrawText(hdc, num.c_str(), -1, &numRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            
            // Draw preview - White text
            RECT previewRect = itemRect;
            previewRect.left += leftOffset + 40;
            previewRect.right -= 180;
            SetTextColor(hdc, RGB(255, 255, 255));
            DrawText(hdc, item->preview.c_str(), -1, &previewRect, 
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            
            // Draw format info - Medium gray
            std::wstring info = item->formatName + L" - " + item->fileType;
            RECT infoRect = itemRect;
            infoRect.left = infoRect.right - 175;
            SetTextColor(hdc, RGB(150, 150, 150));
            
            // Format date
            auto time_t = std::chrono::system_clock::to_time_t(item->timestamp);
            std::wstringstream ss;
            ss << std::put_time(std::localtime(&time_t), L"%H:%M:%S");
            info += L" [" + ss.str() + L"]";
            
            DrawText(hdc, info.c_str(), -1, &infoRect, 
                    DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            
            // Draw separator line - Dark gray
            if (i < endIndex - 1) {
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(40, 40, 40));
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, itemRect.left, itemRect.bottom - 1, nullptr);
                LineTo(hdc, itemRect.right, itemRect.bottom - 1);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
            }
            
            yPos += itemHeight;
        }
        
        EndPaint(hwnd, &ps);
        return 0;
    }
    
    case WM_KEYDOWN:
        // Handle Enter key first
        if (wParam == VK_RETURN) {
            // Enter pressed - paste the selected item or item at typed number
            // Note: Enter in search box is handled by SearchEditProc
            if (!mgr->numberInput.empty()) {
                // Number input takes priority
                try {
                    int itemNumber = std::stoi(mgr->numberInput);
                    // Convert from 1-based to 0-based filtered index
                    int index = itemNumber - 1;
                    if (index >= 0 && index < (int)mgr->filteredIndices.size()) {
                        mgr->PasteItem(index);
                        mgr->numberInput.clear();
                        mgr->HideListWindow();
                    } else {
                        // Invalid number, clear input
                        mgr->numberInput.clear();
                        mgr->UpdateListWindow();
                    }
                } catch (...) {
                    // Invalid number format, clear input
                    mgr->numberInput.clear();
                    mgr->UpdateListWindow();
                }
            } else {
                // No number input - paste selected item
                if (mgr->selectedIndex >= 0 && mgr->selectedIndex < (int)mgr->filteredIndices.size()) {
                    mgr->PasteItem(mgr->selectedIndex);
                    mgr->HideListWindow();
                }
            }
            return 0;
        }
        // Handle Backspace
        if (wParam == VK_BACK) {
            // Backspace - remove last digit
            if (!mgr->numberInput.empty()) {
                mgr->numberInput.pop_back();
                mgr->UpdateListWindow();
            }
            return 0;
        }
        // Quick paste for single digits 1-9 only if no number input is being typed
        // Use virtual key codes VK_0 through VK_9 (0x30-0x39)
        if (wParam >= 0x31 && wParam <= 0x39 && mgr->numberInput.empty()) {
            int index = (wParam - 0x31); // VK_1 (0x31) = item 0, VK_2 (0x32) = item 1, etc.
            if (index < (int)mgr->filteredIndices.size()) {
                mgr->PasteItem(index);
                mgr->HideListWindow();
            }
            return 0;
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
        // Navigate with arrow keys (only if search box doesn't have focus)
        {
            HWND focusedWindow = GetFocus();
            if (focusedWindow != mgr->hwndSearch) {
                if (wParam == VK_UP) {
                    // Move selection up
                    if (mgr->selectedIndex > 0) {
                        mgr->selectedIndex--;
                        // Ensure selected item is visible
                        if (mgr->selectedIndex < mgr->scrollOffset) {
                            mgr->scrollOffset = mgr->selectedIndex;
                        }
                        mgr->UpdateListWindow();
                    }
                    return 0;
                }
                if (wParam == VK_DOWN) {
                    // Move selection down
                    int maxIndex = (int)mgr->filteredIndices.size() - 1;
                    if (mgr->selectedIndex < maxIndex) {
                        mgr->selectedIndex++;
                        // Ensure selected item is visible
                        int maxScroll = std::max(0, (int)mgr->filteredIndices.size() - mgr->itemsPerPage);
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
        if (wParam == VK_PRIOR) { // Page Up
            mgr->scrollOffset = std::max(0, mgr->scrollOffset - mgr->itemsPerPage);
            // Update selection to first visible item
            mgr->selectedIndex = mgr->scrollOffset;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_NEXT) { // Page Down
            int maxScroll = std::max(0, (int)mgr->filteredIndices.size() - mgr->itemsPerPage);
            mgr->scrollOffset = std::min(maxScroll, mgr->scrollOffset + mgr->itemsPerPage);
            // Update selection to first visible item
            mgr->selectedIndex = mgr->scrollOffset;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_HOME) {
            mgr->scrollOffset = 0;
            mgr->selectedIndex = 0;
            mgr->UpdateListWindow();
            return 0;
        }
        if (wParam == VK_END) {
            int maxScroll = std::max(0, (int)mgr->filteredIndices.size() - mgr->itemsPerPage);
            mgr->scrollOffset = maxScroll;
            int maxIndex = (int)mgr->filteredIndices.size() - 1;
            mgr->selectedIndex = maxIndex >= 0 ? maxIndex : 0;
            mgr->UpdateListWindow();
            return 0;
        }
        break;
        
    case WM_CHAR:
        // Handle numeric input (this is called after WM_KEYDOWN)
        if (wParam >= '0' && wParam <= '9') {
            // Limit to 4 digits (enough for 1000 items)
            if (mgr->numberInput.length() < 4) {
                mgr->numberInput += (wchar_t)wParam;
                mgr->UpdateListWindow();
            }
            return 0;
        }
        break;
        
    case WM_VSCROLL:
        {
            // Handle scrollbar messages (even though scrollbar is hidden)
            // This allows programmatic scrolling if needed
            int maxScroll = std::max(0, (int)mgr->filteredIndices.size() - mgr->itemsPerPage);
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
            
            if (itemIndex >= 0 && itemIndex < (int)mgr->filteredIndices.size()) {
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
        
    case WM_LBUTTONDBLCLK:
        {
            // Double-click to paste item
            POINT pt;
            pt.x = LOWORD(lParam);
            pt.y = HIWORD(lParam);
            
            // Calculate which item was clicked
            int clickedIndex = mgr->GetItemAtPosition(pt.x, pt.y);
            
            if (clickedIndex >= 0 && clickedIndex < (int)mgr->filteredIndices.size()) {
                mgr->PasteItem(clickedIndex);
                mgr->numberInput.clear();
                mgr->HidePreviewWindow();
                mgr->HideListWindow();
            }
            return 0;
        }
        
    case WM_COMMAND:
        // Handle search edit control text change
        if (HIWORD(wParam) == EN_CHANGE && (HWND)lParam == mgr->hwndSearch) {
            // Get text from search box
            int length = GetWindowTextLength(mgr->hwndSearch);
            if (length > 0) {
                wchar_t* buffer = new wchar_t[length + 1];
                GetWindowText(mgr->hwndSearch, buffer, length + 1);
                mgr->searchText = buffer;
                delete[] buffer;
            } else {
                mgr->searchText.clear();
            }
            // Filter items based on search text
            mgr->FilterItems();
            return 0;
        }
        break;
        
    case WM_RBUTTONDOWN:
    case WM_CONTEXTMENU:
        {
            // Show context menu on right-click
            POINT pt;
            if (uMsg == WM_RBUTTONDOWN) {
                pt.x = LOWORD(lParam);
                pt.y = HIWORD(lParam);
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
                    ClientToScreen(hwnd, &pt);
                }
            }
            
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, 100, L"Clear List");
            
            SetForegroundWindow(hwnd);
            
            // TrackPopupMenu needs special handling for proper message delivery
            DWORD cmd = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, 
                                       pt.x, pt.y, 0, hwnd, nullptr);
            
            // Post a fake message to ensure menu closes properly
            PostMessage(hwnd, WM_NULL, 0, 0);
            
            if (cmd == 100) {
                mgr->ClearClipboardHistory();
            }
            
            DestroyMenu(hMenu);
            return 0;
        }
        
    case WM_MOUSEWHEEL:
        {
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            int maxScroll = std::max(0, (int)mgr->filteredIndices.size() - mgr->itemsPerPage);
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
            HRGN hRgn = CreateRoundRectRgn(0, 0, rect.right + 1, rect.bottom + 1, 15, 15);
            SetWindowRgn(hwnd, hRgn, TRUE);
            DeleteObject(hRgn);
            break;
        }
    
    case WM_CTLCOLOREDIT:
        {
            // Style the search edit control with dark theme
            HDC hdcEdit = (HDC)wParam;
            SetTextColor(hdcEdit, RGB(255, 255, 255)); // White text
            SetBkColor(hdcEdit, RGB(20, 20, 20)); // Dark gray background
            // Return a brush handle for the background
            static HBRUSH hDarkBrush = CreateSolidBrush(RGB(20, 20, 20));
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
    
    // Handle Enter key in search box
    if (uMsg == WM_KEYDOWN && wParam == VK_RETURN) {
        // Get current search text and update filter before pasting
        int length = GetWindowTextLength(hwnd);
        if (length > 0) {
            wchar_t* buffer = new wchar_t[length + 1];
            GetWindowText(hwnd, buffer, length + 1);
            mgr->searchText = buffer;
            delete[] buffer;
        } else {
            mgr->searchText.clear();
        }
        // Update filter to ensure it's current
        mgr->FilterItems();
        // Paste first filtered item (always index 0)
        if (!mgr->filteredIndices.empty()) {
            mgr->PasteItem(0);
            mgr->HideListWindow();
        }
        return 0;  // Consume the Enter key
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
            L"misc02.ico",
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
        
        // Check if Ctrl is pressed
        bool isCtrlPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        bool isExtended = (pKeyboard->flags & LLKHF_EXTENDED) != 0;
        
        // Detect Ctrl+NumPadDot key down
        // VK_DECIMAL (0x6E) is the NumPad decimal point
        // The extended flag helps distinguish NumPad keys
        if (wParam == WM_KEYDOWN && isCtrlPressed) {
            // Check for NumPad decimal point (either by VK_DECIMAL or extended flag with 0x6E)
            if (pKeyboard->vkCode == VK_DECIMAL || (pKeyboard->vkCode == 0x6E && isExtended)) {
                // Post message to main window to toggle list
                PostMessage(mgr->hwndMain, WM_CLIPBOARD_HOTKEY, 0, 0);
                return 1; // Consume the key so it doesn't propagate
            }
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
    // Initialize filtered indices (show all items initially)
    FilterItems();
    // Reset selection to first item (if any items exist)
    if (!filteredIndices.empty()) {
        selectedIndex = 0;
    } else {
        selectedIndex = -1;
    }
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
    
    if (filteredIndex >= 0 && filteredIndex < (int)filteredIndices.size()) {
        return filteredIndex; // Return filtered index (will be converted to actual index in PasteItem)
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
    int previewWidth = 300;
    int previewHeight = 300;
    
    HBITMAP previewBmp = item->GetPreviewBitmap();
    if (previewBmp) {
        BITMAP bm;
        GetObject(previewBmp, sizeof(BITMAP), &bm);
        int maxSize = 350;
        float scale = std::min((float)maxSize / bm.bmWidth, (float)maxSize / bm.bmHeight);
        previewWidth = (int)(bm.bmWidth * scale) + 20;
        previewHeight = (int)(bm.bmHeight * scale) + 20;
        previewWidth = std::min(previewWidth, 450);
        previewHeight = std::min(previewHeight, 450);
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
    
    // Update sequence number to prevent detecting our own clipboard change
    lastSequenceNumber = GetClipboardSequenceNumber();
    
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
            // Update sequence number after setting clipboard
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
            
            // Clear flag after a short delay to allow paste to complete
            Sleep(100);
        }
        
        isPasting = false;
    } else {
        isPasting = false;
    }
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
                    // Create clipboard item with primary format (or converted format)
                    std::unique_ptr<ClipboardItem> item;
                    try {
                        // Validate data size before creating item
                        if (data.size() > 0 && data.size() < 200 * 1024 * 1024) { // Limit to 200MB
                            item = std::make_unique<ClipboardItem>(primaryFormat, data);
                            if (!item) {
                                CloseClipboard();
                                isProcessingClipboard = false;
                                return;
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
                    
                    // If we converted CF_BITMAP to DIB, store it as DIB
                    if (primaryFormat == CF_BITMAP && storageFormat == CF_DIB) {
                        item->AddFormat(CF_DIB, data);
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
                    
                    // Only add if not a duplicate
                    if (!isDuplicate && item) {
                        try {
                            // Add to front of list (most recent first)
                            clipboardHistory.insert(
                                clipboardHistory.begin(),
                                std::move(item)
                            );
                            
                            // Limit to MAX_ITEMS - aggressively clean up old items to prevent memory leaks
                            while (clipboardHistory.size() >= MAX_ITEMS) {
                                // Remove oldest item (at the back) - destructor will clean up bitmaps
                                clipboardHistory.pop_back();
                            }
                            
                            // Also limit memory usage - simplified calculation to prevent crashes
                            // If we have more than 50 items, remove oldest until we're at 50
                            const int MEMORY_SAFE_LIMIT = 50;
                            while (clipboardHistory.size() > MEMORY_SAFE_LIMIT) {
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

void ClipboardManager::PlayClickSound() {
    // Simple, fast, reliable sound playback using mciSendString
    // This plays the MP3 file from the executable directory
    
    static std::wstring soundPath;
    static bool initialized = false;
    
    if (!initialized) {
        initialized = true;
        // Get executable directory once
        wchar_t exePath[MAX_PATH];
        GetModuleFileNameW(nullptr, exePath, MAX_PATH);
        wchar_t* lastSlash = wcsrchr(exePath, L'\\');
        if (lastSlash) *(lastSlash + 1) = L'\0';
        soundPath = std::wstring(exePath) + L"click.mp3";
    }
    
    // Stop any previous playback
    mciSendString(L"stop clicksnd", nullptr, 0, nullptr);
    mciSendString(L"close clicksnd", nullptr, 0, nullptr);
    
    // Open with MPEGVideo device type (required for MP3)
    std::wstring openCmd = L"open \"" + soundPath + L"\" type MPEGVideo alias clicksnd";
    if (mciSendString(openCmd.c_str(), nullptr, 0, nullptr) == 0) {
        // Play asynchronously
        mciSendString(L"play clicksnd", nullptr, 0, nullptr);
    }
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

void ClipboardManager::FilterItems() {
    filteredIndices.clear();
    
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
            
            // Search in preview, format name, and file type
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

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    ClipboardManager manager;
    
    if (!manager.Initialize()) {
        MessageBox(nullptr, L"Failed to initialize clip2", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    manager.Run();
    
    return 0;
}

