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
#include <array>

struct ClipboardItem {
    UINT format;  // Primary format (for display/compatibility)
    mutable std::map<UINT, std::vector<BYTE>> formats;  // All formats (mutable: lazy blob hydration)
    std::wstring preview;
    std::wstring formatName;
    std::wstring fileType;
    std::chrono::system_clock::time_point timestamp;
    HBITMAP thumbnail;
    HBITMAP previewBitmap;  // Larger preview for hover
    bool thumbnailAttempted;   // Lazy: tried building the 48x48 thumbnail already
    bool previewAttempted;     // Lazy: tried building the hover preview already
    bool isImage;
    bool isVideo;
    bool pinned;  // Pinned/favorite items stay on top, survive Clear, and always persist
    
    ClipboardItem(UINT fmt, const std::vector<BYTE>& d) 
        : format(fmt), timestamp(std::chrono::system_clock::now()), thumbnail(nullptr), previewBitmap(nullptr), thumbnailAttempted(false), previewAttempted(false), isImage(false), isVideo(false), pinned(false), formatName(L"Unknown Format"), fileType(L"Other"), preview(L"[Unknown]"), searchIndexDirty(false) {
        // CRITICAL: Store data first, before any processing
        if (d.empty() || d.size() > 200 * 1024 * 1024) {
            return;
        }
        
        // Store format data immediately - this must succeed
        formats[fmt] = d;
        
        InitFormatMetadata(fmt);

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
        
        // Thumbnails are built lazily (EnsureThumbnail) the first time a row needs one,
        // so captures never pay for image decoding up front.

        RebuildSearchIndex();
    }
    
    // Primary format is a blob still sitting on disk (history load). Metadata is derived
    // from the format id alone -- only text previews read the payload, and text formats
    // are never deferred. Delegates to the byte constructor with an empty payload, whose
    // early-out leaves `formats` empty, which is exactly what a deferred primary wants.
    struct DeferredPrimary {};
    ClipboardItem(UINT fmt, DeferredPrimary)
        : ClipboardItem(fmt, std::vector<BYTE>()) {
        InitFormatMetadata(fmt);
        preview = L"[" + formatName + L"]";
        RebuildSearchIndex();
    }

    // Add additional format.
    // Rebuilding the search index is a full pass over up to 500KB of text plus a
    // trigram pass, so callers that add SEVERAL formats in a row (clipboard capture,
    // history load) pass deferIndex=true and call FinalizeSearchIndex() once at the
    // end -- otherwise an item with text + RTF + HTML reindexes itself three times.
    // The rvalue overload avoids copying the payload a second time into the map.
    void AddFormat(UINT fmt, const std::vector<BYTE>& d, bool deferIndex = false) {
        formats[fmt] = d;
        pendingBlobs.erase(fmt);   // explicit bytes supersede any on-disk reference
        blobOrigin.erase(fmt);
        NoteTextFormatAdded(fmt, deferIndex);
    }
    void AddFormat(UINT fmt, std::vector<BYTE>&& d, bool deferIndex = false) {
        formats[fmt] = std::move(d);
        pendingBlobs.erase(fmt);
        blobOrigin.erase(fmt);
        NoteTextFormatAdded(fmt, deferIndex);
    }
    // Rebuild the index if any deferred AddFormat touched a text format. Cheap no-op otherwise.
    void FinalizeSearchIndex() {
        if (searchIndexDirty) RebuildSearchIndex();
    }
    
    // ---- Blob-backed formats (lazy load) ----
    // Formats >= 256KB are persisted as content-addressed sidecar files under
    // %APPDATA%\clip2\blobs. Reading every one at startup put the entire image history
    // in RAM before the tray icon appeared, so non-text blobs are recorded here as a
    // digest and read on first real use. Because the sidecar IS the backing store,
    // hydrated bytes can also be handed back (DehydrateBlobFormats) and re-read later.
    // Text blobs are never deferred -- RebuildSearchIndex needs their bytes at load time.
    typedef std::array<BYTE, 20> BlobDigest;
    mutable std::map<UINT, BlobDigest> pendingBlobs;  // format -> digest, bytes not read yet
    mutable std::map<UINT, BlobDigest> blobOrigin;    // format -> digest for bytes we DID read

    void AddPendingBlob(UINT fmt, const BYTE digest[20]) {
        BlobDigest d;
        memcpy(d.data(), digest, 20);
        pendingBlobs[fmt] = d;
        formats.erase(fmt);
    }
    // True when the item carries a format at all, loaded or still on disk. Callers that
    // used to test formats.empty() must use this, or a deferred image looks like nothing.
    bool HasAnyFormat() const { return !formats.empty() || !pendingBlobs.empty(); }
    bool HasFormat(UINT fmt) const {
        return formats.count(fmt) != 0 || pendingBlobs.count(fmt) != 0;
    }
    size_t FormatCount() const { return formats.size() + pendingBlobs.size(); }
    // Read every outstanding blob. Required before iterating `formats` for its bytes.
    void HydrateAllFormats() const;
    // Release bytes that can be re-read from a sidecar, turning them back into pending
    // references. Text is left alone (the search index reads it).
    void DehydrateBlobFormats();
    // Drop a format entirely, loaded or pending.
    void RemoveFormat(UINT fmt) {
        formats.erase(fmt);
        pendingBlobs.erase(fmt);
        blobOrigin.erase(fmt);
    }

    // Get data for a specific format. Hydrates from the sidecar on demand.
    const std::vector<BYTE>* GetFormatData(UINT fmt) const {
        auto it = formats.find(fmt);
        if (it != formats.end()) {
            return &it->second;
        }
        if (pendingBlobs.count(fmt) != 0 && HydrateFormat(fmt)) {
            it = formats.find(fmt);
            if (it != formats.end()) return &it->second;
        }
        return nullptr;
    }
    
    // Get full text content for search (handles line breaks; empty for non-text)
    std::wstring GetFullSearchableText() const;

    // ---- Search index (built once per item, not per keystroke) ----
    std::wstring searchPreviewLower;            // lowercase preview + formatName + fileType
    std::wstring searchBodyLower;               // lowercase full searchable text (<= 500KB)
    std::vector<unsigned char> trigramBloom;    // 1KB bloom filter of text trigrams (no false negatives)
    std::vector<unsigned char> charBloom;       // 32B presence bitmap of text chars (gates 1-2 char needles)
    bool searchIndexDirty;                      // a deferred AddFormat changed body text
    void RebuildSearchIndex();                  // call after any text-format mutation
    
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
    
    HBITMAP GetPreviewBitmap();  // Generate larger hover preview on demand (lazy)
    void EnsureThumbnail();      // Generate the 48x48 thumbnail on first need (lazy)

    // ---- Cached bitmap lifetime ----
    // Both cached bitmaps used to live until the item itself died, so browsing a long
    // image history grew the process without bound: hover previews are up to 500x500
    // (~1MB each) and every thumbnail pins one GDI object. They are now released under
    // a cap (ClipboardManager::TrimBitmapCaches) and regenerate lazily on next use.
    // useTick orders them by recency; 0 means "never used".
    unsigned long long previewUseTick = 0;
    unsigned long long thumbUseTick = 0;
    void ReleasePreviewBitmap() {
        if (previewBitmap) { DeleteObject(previewBitmap); previewBitmap = nullptr; }
        previewAttempted = false;   // allow regeneration on next hover
        previewUseTick = 0;
    }
    void ReleaseThumbnail() {
        if (thumbnail) { DeleteObject(thumbnail); thumbnail = nullptr; }
        thumbnailAttempted = false;
        thumbUseTick = 0;
    }

private:
    // Shared tail of both AddFormat overloads: only text formats affect the index.
    // Read one outstanding blob into `formats`. Defined in the .cpp beside the blob
    // directory helpers. Returns false when the sidecar is missing or unreadable.
    bool HydrateFormat(UINT fmt) const;

    void NoteTextFormatAdded(UINT fmt, bool deferIndex) {
        if (fmt != CF_TEXT && fmt != CF_UNICODETEXT && fmt != CF_OEMTEXT) return;
        if (deferIndex) searchIndexDirty = true;
        else RebuildSearchIndex();
    }

    // Format name / file type / isImage, all derived from the format id alone.
    void InitFormatMetadata(UINT fmt) {
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
        
    }

    std::wstring GetFormatName(UINT fmt);
    std::wstring GetFileType(UINT fmt);
    std::wstring GetPreview(const std::vector<BYTE>& data, UINT fmt);
    void GenerateThumbnail();
    void GeneratePreviewBitmap();  // large hover preview (<= 500x500), DIB items only
    HBITMAP CreateBitmapFromData();
};

struct MergedRichPayload;

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
    // Release cached thumbnails / hover previews beyond a recency cap so a long image
    // history cannot grow the process without bound. Cheap: one pass over the history.
    void TrimBitmapCaches();
    int GetItemAtPosition(int x, int y);
    void PasteItem(int index);
    void PasteMultipleItems();
    // Excel special paste (Z): for each multi-selected item, F2 (edit cell),
    // Ctrl+V (rich formats), then Enter to commit and move down one row.
    void PasteExcelSelection();
    // Merge the current multi-selection into one new history item at the top.
    // plainOnly=false keeps RTF/HTML; plainOnly=true stores Unicode text only.
    void MergeSelectedItems(bool plainOnly = false);
    // P + multi-select: paste selection as plain text (one line per item) and also
    // insert that plain merged text as a new top history item.
    void PastePlainAndMergeSelection();
    // P + single item: paste that item as plain Unicode (no RTF/HTML).
    bool PasteItemAsPlainText(int filteredIndex);
    // Insert a previously built merge payload as a new unpinned history item at the top.
    // Does not touch the system clipboard. Returns false if unicode is empty / insert fails.
    bool InsertMergedPayloadIntoHistory(const MergedRichPayload& merged, bool plainOnly);
    // Smart paste: put a derived plain-text payload on the clipboard and Ctrl+V it.
    void PasteTransformedText(const std::wstring& text);
    // One-key smart paste modes (SMART_PASTE_*). Returns false (silently) when not applicable.
    bool SmartPasteItem(int filteredIndex, int mode);
    // Modal edit dialog. saveAsNew=false (E): edit then paste. saveAsNew=true (X): save edits as a new history item.
    void ShowEditPasteDialog(bool saveAsNew = false);
    static LRESULT CALLBACK EditPasteDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK EditPasteEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
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
    void MarkHistoryDirty();   // Debounced persistence: schedules a save ~1.5s after the last mutation
    void FilterItems();
    void FilterSnippets();
    // While typing a paste number (#N), temporarily move that item to the top of the
    // active list so it can be previewed / acted on. Restores natural order when cleared.
    void ApplyNumberInputPromotion();
    int OriginalItemNumber(int displayFilteredIndex) const;  // 1-based label from unpromoted order
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
    bool historyDirty;        // History changed since the last save (debounced via TIMER_SAVE_HISTORY)
    static const UINT_PTR TIMER_SAVE_HISTORY = 2;  // hwndMain timer id (1 = overlay focus check)
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
    std::vector<int> filteredIndicesBase;  // Search/history order (before number-input promotion)
    std::vector<int> filteredIndices;      // Display order (may temporarily promote a typed #)
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
    HWND hwndEditPaste; // Edit dialog window (paste mode or save-as-new mode)
    bool editPasteSaveAsNew; // True when the edit dialog should commit a new history item (X)
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
    // Quick add (editIndex < 0) or edit (actual snippets[] index) popup from the snippets overlay.
    void ShowSnippetEditorDialog(int editIndex = -1);
    
    // Bypass copy blocks: capture clipboard immediately when it changes so we have content before apps clear it
    bool TryCaptureClipboardImmediately();
    void ProcessClipboardFromSnapshot();
    
    static LRESULT CALLBACK SnippetsManagerProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SnippetEditorDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK SnippetEditorEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    
    std::vector<Snippet> snippets;
    HWND hwndSnippetsManager;
    HWND hwndSnippetEditor;       // Quick add/edit snippet dialog
    int snippetEditorEditIndex;   // -1 = add new; otherwise index into snippets[]
    bool ignoreNextSnippetShortcutChar; // Consume WM_CHAR for A/E after the shortcut opens the dialog
    
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
    static constexpr int DEFAULT_MAX_ITEMS = 300;   // Default history cap when no user override exists
    static constexpr int MIN_MAX_ITEMS = 10;        // Lower clamp for the configurable history size
    static constexpr int MAX_MAX_ITEMS = 2000;      // Upper clamp for the configurable history size
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

    // Smart paste modes (also used as context-menu command IDs)
    enum SmartPasteMode {
        SMART_PASTE_URL_CLEAN = 120,   // U: strip tracking params from URL, paste
        SMART_PASTE_MARKDOWN = 121,    // M: paste as Markdown link
        SMART_PASTE_FILEPATH = 122,    // File path (context menu; P is now plain text)
        SMART_PASTE_HTML_PLAIN = 123,  // H: HTML -> plain text, paste
        SMART_PASTE_EDIT = 124,        // E: edit/merge before paste (menu entry)
        SMART_PASTE_EDIT_SAVE = 125,   // X: edit and save as a new history item
        SMART_MERGE_SELECTION = 126,   // M (with multi-select): merge into new top item
        SMART_PASTE_EXCEL = 127,       // Z (with multi-select): Excel F2 + paste + Enter per item
        SMART_PASTE_PLAIN_MERGE = 128, // P (with multi-select): paste + merge as plain text
        SMART_PASTE_PLAIN = 129        // P (single item): paste as plain text
    };
};

