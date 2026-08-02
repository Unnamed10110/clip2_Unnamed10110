# clip2

A minimalistic and superfast clipboard manager for Windows with OLED theme and useful features.


## Features

- ✅ **Background Operation**: Runs silently in the system tray
- ✅ **Universal Clipboard Support**: Captures text, rich text, images, files, and more
- ✅ **LIFO History**: Most recently copied item appears first (up to 1000 items)
- ✅ **Hotkey Access**: Press **Ctrl+NumPadDot** to show/hide the clipboard list
- ✅ **Number Paste**: Type an item number to bring it to the top of the list in real time, then **Enter** to paste (or interact with it first — preview, right-click, smart paste, etc.)
- ✅ **Visual Feedback**: Click sound plays when items are copied
- ✅ **System Tray Icon**: Always visible when running (uses ico2.ico)
- ✅ **AMOLED Neon Themes**: Pure-black background with neon accents — switch between Neon Green (AS/400 5250 default), Red, Blue, Cyan, Purple, Yellow, Orange, and White from the Settings dialog
- ✅ **Custom Font Color & Family**: Settings lets you override the overlay font color with the standard Windows color picker (any RGB value), and pick the overlay/search font from every font family installed on the system. Both are persisted and applied live
- ✅ **Full Per-Element Colors**: Settings → "Element colors" exposes a swatch for every overlay color — Background, Text, Accent (selection), Selected text, Border, and Dim/secondary. Click a swatch to pick any RGB value, double-click to reset that element to the active preset. Changes apply live and persist (registry `HKEY_CURRENT_USER\Software\clip2\Color*`); "Defaults" clears them all
- ✅ **Search Functionality**: Deep search across full content (Ctrl+F), including text after line breaks
- ✅ **Image/Video Thumbnails**: Visual previews for images and videos
- ✅ **Format Preservation**: Pastes with original formatting (bold, colors, etc.)
- ✅ **Plain Text Mode**: Hold Ctrl while pasting for plain text
- ✅ **Start with Windows**: Option to auto-start with Windows session
- ✅ **Duplicate Prevention**: Automatically skips identical consecutive copies
- ✅ **Delete Items**: Remove specific items from history without clearing everything
- ✅ **Multi-Paste**: Select and paste multiple items at once, each on a new line
- ✅ **Text Transformations**: Transform text items (uppercase, lowercase, title case, remove formatting, etc.)
- ✅ **Customizable Hotkeys**: Change the default hotkey through Settings menu
- ✅ **Movable Window**: Drag the list window to position it anywhere on screen
- ✅ **Smart Paste Detection**: Prevents pasted items from being re-added to history
- ✅ **Snippets/Templates**: Predefined text templates with placeholders; supports Rich Text (RTF)
- ✅ **Pinned/Favorite Items**: Right-click → **Pin** to add an item to the left pinned panel; in the main list they keep their normal history position. Pinned items survive **Clear List** and are always persisted (shown with a left accent bar)
- ✅ **Full Persistence**: History now persists images and files (not just text) across restarts
- ✅ **Encrypted History**: `history.dat` is encrypted at rest with Windows DPAPI (per-user); legacy plaintext files are migrated automatically
- ✅ **Configurable History Size**: Set the history cap in Settings (10–2000 items)
- ✅ **Fuzzy Search**: Subsequence matching with ranking (best matches first) in clipboard and snippets overlays
- ✅ **Quick Actions**: Right-click an item for **Copy as plain text**, **Open URL** (for http/https items), and **Save image to file...**
- ✅ **Clipboard Persistence**: History saved to disk; survives app restarts
- ✅ **Fast Performance**: Optimized for minimal CPU usage and memory footprint
- ✅ **Snippets Overlay**: **Ctrl+Right** = Snippets, **Ctrl+Left** = Clipboard (when overlay is open)
- ✅ **Copy from focused control**: When an app never puts content on the clipboard, use tray → **Copy from focused control** to read text from the focused control via UI Automation and add it to history (requires build with UIAutomation; see Technical Details)
- ✅ **Smart Paste Modes**: With the overlay open, one key pastes a derived version of the selected item — **U** = URL with tracking params stripped (`utm_*`, `fbclid`, `gclid`, ...), **M** = Markdown link (`[title](url)` for URLs, `[name](path)` for a single file), **P** = file path(s) as text, **H** = HTML converted to plain text. Also available via right-click → **Paste as**. Keys fail silently when the item doesn't fit the mode
- ✅ **Edit / Merge Before Paste**: Press **E** (or right-click → **Edit & Paste...**) to open a small editor prefilled with the selected item — or with all multi-selected items joined line-by-line — tweak the text, then **Ctrl+Enter** or the **Paste** button pastes the edited result (history is not modified)
- ✅ **Edit & Save as New**: Press **X** (or right-click → **Edit & Save as new...**) to edit the selection the same way, then **Save** / **Ctrl+Enter** adds the result as a new history item at the top (original unchanged) and puts it on the system clipboard
- ✅ **Debounced Incremental Saves**: History is saved ~1.5 s after the last change (copy, delete, pin, transform, clear) instead of only at exit — mutations can no longer be lost to a crash. Writes are atomic (temp file + rename)
- ✅ **Blob Sidecar Storage**: Clipboard formats ≥ 256 KB (large images, file lists) are stored once as content-addressed files under `%APPDATA%\clip2\blobs\`, so pinning or deleting a text item no longer rewrites megabytes of unchanged image data; unreferenced blobs are garbage-collected after each save
- ✅ **Indexed Search**: Each item caches a lowercase copy of its text plus a trigram bloom filter; typing in the search box scores only candidate items instead of re-extracting and re-lowercasing every item's full text per keystroke — search stays instant with hundreds of large items
- ✅ **Lazy Thumbnails & Previews**: Image thumbnails are decoded only when their row first becomes visible in the active panel, and the large hover preview only on first hover; the pinned-panel snapshot draws lightweight type badges, so painting never decodes images

## Building

### Prerequisites

- Windows 10/11
- CMake 3.15 or later
- C++ compiler (MSVC, MinGW, or Clang)
- Visual Studio 2019/2022 (optional, for MSVC)

### Build Steps

1. **Using CMake (Recommended)**:
   ```bash
   mkdir build
   cd build
   cmake .. -G "Visual Studio 17 2022" -A x64
   cmake --build . --config Release
   ```

2. **Using the batch script**:
   ```bash
   build.bat
   ```

3. **Manual compilation** (MinGW):
   ```bash
   windres ClipboardManager.rc -o ClipboardManager.res
   g++ -std=c++17 -O2 -mwindows ClipboardManager.cpp ClipboardManager.res -o ClipboardManager.exe -luser32 -lgdi32 -lshell32 -lwinmm -lcrypt32
   ```

The executable will be created as `clip2.exe`.

## Usage

1. **Start the application**: Run `clip2.exe`
   - The icon will appear in the system tray

2. **View clipboard history**: Press **Ctrl+NumPadDot**
   - A floating window shows your clipboard history with modern OLED theme
   - Each item displays: number, preview, format type, and timestamp
   - Images and videos show thumbnails

3. **Paste an item**: 
   - Type an item number (e.g. `12`) — that item jumps to the top so you can preview or act on it; press **Enter** to paste
   - Use **Arrow keys** to navigate and **Enter** to paste the selected item
   - **Double-click** an item to paste it
   - Hold **Ctrl** while pasting for plain text mode
   - **Smart paste**: **U** clean URL, **M** Markdown link, **P** file path, **H** HTML → plain text (also right-click → **Paste as**)
   - **Edit before paste**: **E** opens an editor with the selection (or the joined multi-selection); **Ctrl+Enter** pastes the edited text
   - **Edit & save as new**: **X** opens the same editor; **Ctrl+Enter** / **Save** keeps the edits as a new clipboard item (original stays)

4. **Search**: Press **Ctrl+F** to focus the search box and filter items
   - Search looks through the **full content** of each item, including text after line breaks
   - Works in both clipboard and snippets overlays

5. **Hide the list**: Press **Ctrl+NumPadDot** again, **Escape**, or click outside the window

6. **System tray menu**: Right-click the tray icon for options
   - **Show Clipboard**: Show the clipboard list
   - **Snippets**: Submenu to paste predefined text templates (or **Snippets...** to manage)
   - **Start with Windows**: Toggle auto-start with Windows (checkmark indicates enabled)
   - **Settings**: Configure hotkeys and other settings
   - **Exit**: Close the application

7. **Delete items**: Press **Delete** key or right-click → **Delete** to remove specific items from history

8. **Multi-paste**: Select multiple items with **Ctrl+Click** or **Shift+Arrow keys**, then press **Enter** to paste them sequentially (each item on a new line)

9. **Text transformations**: Right-click on text items → **Transform Text** for options:
   - Uppercase / Lowercase / Title Case
   - Remove line breaks / Trim whitespace
   - Remove formatting (paste as plain text)

10. **Move window**: Click and drag the title bar or empty areas to move the list window

11. **Snippets/templates**:
    - **Tray menu**: Right-click tray icon → **Snippets** → select a snippet to paste
    - **Snippets overlay**: With the clipboard list open, press **Ctrl+Right** to switch to snippets view; **Ctrl+Left** to switch back to clipboard
    - Search, select (arrows/numbers), and paste snippets the same way as clipboard items
    - **Single-click** on a snippet to paste it and close the overlay
    - **\*set shortcut**: In snippets overlay, type `*set` and press **Enter** to open Manage Snippets (cannot be used as a snippet name)
    - **Placeholders** (in snippet content): `{{date}}`, `{{time}}`, `{{datetime}}`, `{{year}}`, `{{month}}`, `{{day}}`, `{{hour}}`, `{{minute}}`, `{{second}}`, `{{clipboard}}`
    - **Manage Snippets**: Add, edit, delete snippets; Rich Edit supports RTF formatting (bold, colors, etc.)
    - **Quick add / edit**: In the snippets overlay (list focused, not the search box), press **A** to add a snippet or **E** to edit the selected one; also available from the right-click menu

## Technical Details

- **Maximum History**: configurable (default 300, range 10–2000) via Settings; stored in `HKEY_CURRENT_USER\Software\clip2\MaxItems`
- **Supported Formats**: Text, Unicode text, Rich Text Format (RTF), HTML, images (bitmap, DIB, DIBV5), files, and more
- **Format Preservation**: All clipboard formats are preserved when pasting (bold, colors, sizes, etc.)
- **Memory Limits**: 
  - Max 5 formats per item
  - Max 5MB per individual format
  - Max 10MB total per clipboard item
  - Cleanup when exceeding 300 items
- **Hotkey**: Ctrl+NumPadDot (low-level keyboard hook for reliable detection)
- **Update Frequency**: Asynchronous clipboard monitoring with retry logic
- **Duplicate Detection**: Prevents adding identical consecutive items by comparing all formats and data
- **Sound**: Embedded click.mp3 plays on copy (fallback to file in exe directory); MCI warm-up ensures first copy plays sound
- **Navigation**: Arrow keys, Page Up/Down, Home/End for list navigation
- **Auto-hide**: List automatically hides when focus moves to another window
- **Clipboard persistence**: Saved to `%APPDATA%\clip2\history.dat`. Encrypted at rest with Windows DPAPI (container magic `CLP3`); decrypted payload stores all formats per item (text, images, files) plus the pinned flag (`CLP2` version 2). Pinned items are always saved (exempt from the count cap); other items are saved up to the configured history size. Legacy `CLP2` v1 (text-only, plaintext) files load and are upgraded to the encrypted format on next save. Saves are debounced (~1.5 s after the last history change), flushed at exit, and written atomically (`history.dat.tmp` + rename). Formats ≥ 256 KB live as content-addressed sidecar files in `%APPDATA%\clip2\blobs\<sha1>` referenced from the encrypted index; unreferenced blobs are deleted after each successful save
- **Snippets storage**: Registry `HKEY_CURRENT_USER\Software\clip2\Snippets`
- **Search**: Full-text search up to 500KB per item; finds matches across line breaks. Backed by a per-item index (cached lowercase text + trigram bloom filter) so only candidate items are fuzzy-scored per keystroke
- **Copy from focused control**: Uses Windows UI Automation to read selected (or full) text from the focused control when the app does not put anything on the clipboard. The tray menu item appears only if the project is built with the UIAutomation library (e.g. when the Windows SDK is available and CMake finds `UIAutomation`). If you build with MinGW and the menu item is missing, build with Visual Studio and the Windows SDK to enable it.

## Requirements

- Windows API libraries: `user32`, `gdi32`, `shell32`, `winmm`, `comctl32`, `shlwapi`, `ole32`, `crypt32`
- C++17 standard
- ico2.ico icon file (included)
- click.mp3 sound file (embedded as resource, with file fallback)

## Troubleshooting

- **Icon not showing**: Ensure `ico2.ico` is in the same directory or compiled as a resource
- **Hotkey not working**: Ensure NumLock is enabled or try the extended key combination
- **Sound not playing**: Ensure `click.mp3` is in the same directory as `clip2.exe` (or embedded as resource)
- **Performance issues**: The application uses minimal resources; if issues occur, check system resources

## Changelog

### Theme Font Controls
- ✅ **Custom overlay font color**: Settings → "Pick color..." opens the standard Windows color dialog, lets you choose any RGB value, and applies it live to the clipboard/snippets overlay (text + selection accent). Stored under `HKEY_CURRENT_USER\Software\clip2\ThemeFontColor`; clearing the override via Defaults restores the active preset's neon accent.
- ✅ **All-system-font picker**: Settings → "Overlay font" lists every font family installed on the machine (enumerated via `EnumFontFamiliesExW`, de-duplicated, vertical `@`-prefixed faces skipped). Selected face is applied to both the overlay list and the search box and persisted under `HKEY_CURRENT_USER\Software\clip2\ThemeFontFace`. Default remains `Consolas` for the terminal look.

### Latest Reliability & Theming Pass
- ✅ **Reworked paste pipeline**: Single-item, multi-item, snippet, and keystroke paste now share a small set of helpers (clipboard transactions with retry, plain-text extraction, RTF byte serialization, focus restoration). This eliminates timing races between fragmented copies of the same logic.
- ✅ **Reliable multi-item paste**: When every selected item is text/RTF/HTML, the manager now builds a single combined payload (Unicode + optional RTF wrapper) and sends one `Ctrl+V`, instead of rapidly swapping the clipboard between events. Items containing images or other binary formats still replay sequentially, but with longer settle times so trailing items never get dropped.
- ✅ **Rich snippet paste**: RTF snippets keep their formatting in rich-text-aware targets (Word, Outlook, OneNote, browsers), while a clean Unicode fallback is published for plain-text targets. Legacy snippets without a stored plain version now derive plain text via a hidden Rich Edit instead of leaking RTF source as plain text.
- ✅ **Robust keystroke paste**: The synthetic Unicode typing path uses smaller batches (12 chars), retries partial `SendInput` results from the next index instead of dropping the tail, converts `\r\n`/`\n`/`\t` to real `VK_RETURN`/`VK_TAB` events, and paces itself so long content arrives complete in slower editors.
- ✅ **AMOLED neon theme switcher**: Settings → Overlay theme lets you live-preview and persist Neon Green / Red / Blue / Cyan / Purple / Yellow / Orange / White presets. Selection is stored in `HKEY_CURRENT_USER\Software\clip2\ThemeId` and applied at startup.

### Recent Features
- ✅ **Clipboard persistence**: History saved to `%APPDATA%\clip2\history.dat`; survives app restarts (up to 300 items)
- ✅ **Snippets overlay**: **Ctrl+Right** / **Ctrl+Left** to switch between clipboard and snippets; single-click to paste snippets
- ✅ **Rich Text snippets**: Manage Snippets uses Rich Edit; RTF content pasted with formatting preserved
- ✅ **\*set shortcut**: Type `*set` + Enter in snippets overlay to open Manage Snippets
- ✅ **Deep search**: Search filters by full content (including text after line breaks), not just preview
- ✅ **Delete individual items**: Delete key or right-click context menu to remove specific items
- ✅ **Multi-paste mode**: Select multiple items (Ctrl+Click or Shift+Arrow) and paste them sequentially with newlines
- ✅ **Text transformations**: Right-click menu for text manipulation (uppercase, lowercase, title case, remove line breaks, trim whitespace, remove formatting)
- ✅ **Customizable hotkeys**: Settings menu to change the toggle hotkey (default: Ctrl+NumPadDot)
- ✅ **Movable window**: Drag the list window by clicking on the title bar or empty areas
- ✅ **Smart paste detection**: Prevents pasted items from being re-added to clipboard history

### Recent Improvements & Fixes
- ✅ **First-copy sound**: MCI warm-up at startup so click sound plays on first Ctrl+C
- ✅ **Crash reporting**: Unhandled exceptions show error dialog instead of silent exit
- ✅ **Clipboard file I/O**: Uses Windows API (CreateFileW/ReadFile) for reliable history load/save on MinGW

### Stability Enhancements
- ✅ **Fixed crashes with multiple images**: Properly handles CF_BITMAP format conversion to DIB
- ✅ **Fixed crashes with rich text from Word**: Added extensive validation and error handling for RTF and complex formats
- ✅ **Fixed crashes with rapid copying**: Implemented re-entrant call protection with `isProcessingClipboard` flag
- ✅ **Fixed memory leaks**: Aggressive memory management with strict limits per format and item
- ✅ **Fixed long-running stability**: Proper COM initialization handling and resource cleanup

### Memory Management
- Reduced maximum history from 1000 to 100 items for better stability
- Limited to 5 formats per clipboard item (down from unlimited)
- Max 5MB per individual format (down from 100MB)
- Max 10MB total per clipboard item
- Automatic cleanup when history exceeds 300 items
- Forced memory deallocation on history clear
- Simplified preview generation to prevent resource exhaustion

### Error Handling
- Extensive try-catch blocks around clipboard operations
- Validation of all data sizes, pointers, and structures before processing
- Safe bounds checking for text and binary data
- Proper cleanup in all error paths (clipboard, GDI resources, COM objects)
- Prevention of buffer overruns and invalid memory access
- **Robust clipboard failure handling**:
  - Retry logic: up to 5 retries with exponential backoff (50ms, 100ms, 150ms, etc.)
  - Quick retries: 10 quick attempts (5ms delays) before giving up
  - Graceful degradation: if clipboard can't be opened, skips that item and continues
  - Prevents crashes from temporary clipboard locks or unavailability

### Performance Optimizations
- Streamlined clipboard format storage and retrieval
- Optimized duplicate detection for large formats
- Reduced preview bitmap complexity
- Minimized GDI object creation and destruction
- Efficient filtering and search operations

## Developer Information

- **Developer**: Unnamed10110
- **Email**: trojan.v6@gmail.com
- **Alternative Email**: sergiobritos10110@gmail.com

## License

This project is provided as-is for personal use.

