# clip2

A minimalistic and superfast clipboard manager for Windows with OLED theme and useful features.


## Features

- ✅ **Background Operation**: Runs silently in the system tray
- ✅ **Universal Clipboard Support**: Captures text, rich text, images, files, and more
- ✅ **LIFO History**: Most recently copied item appears first (up to 1000 items)
- ✅ **Hotkey Access**: Press **Ctrl+NumPadDot** to show/hide the clipboard list
- ✅ **Quick Paste**: Press **1-9** or type number + Enter to paste items from the list
- ✅ **Visual Feedback**: Click sound plays when items are copied
- ✅ **System Tray Icon**: Always visible when running (uses ico2.ico)
- ✅ **AS/400 5250 Theme**: Phosphor green on black, monospace font, sharp corners—classic terminal emulator look
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
- ✅ **Clipboard Persistence**: History saved to disk; survives app restarts (up to 50 text items)
- ✅ **Fast Performance**: Optimized for minimal CPU usage and memory footprint
- ✅ **Snippets Overlay**: Switch between clipboard and snippets view with SS (press S twice)

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
   g++ -std=c++17 -O2 -mwindows ClipboardManager.cpp ClipboardManager.res -o ClipboardManager.exe -luser32 -lgdi32 -lshell32 -lwinmm
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
   - With the list visible, press **1-9** for quick paste, or type number + **Enter** for multi-digit selection
   - Use **Arrow keys** to navigate and **Enter** to paste selected item
   - **Double-click** an item to paste it
   - Hold **Ctrl** while pasting for plain text mode

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
    - **Snippets overlay**: With the clipboard list open, press **S** twice (**SS**) to switch to snippets view; press **SS** again to switch back to clipboard
    - Search, select (arrows/numbers), and paste snippets the same way as clipboard items
    - **Single-click** on a snippet to paste it and close the overlay
    - **\*set shortcut**: In snippets overlay, type `*set` and press **Enter** to open Manage Snippets (cannot be used as a snippet name)
    - **Placeholders** (in snippet content): `{{date}}`, `{{time}}`, `{{datetime}}`, `{{year}}`, `{{month}}`, `{{day}}`, `{{hour}}`, `{{minute}}`, `{{second}}`, `{{clipboard}}`
    - **Manage Snippets**: Add, edit, delete snippets; Rich Edit supports RTF formatting (bold, colors, etc.)

## Technical Details

- **Maximum History**: 100 items (optimized for stability)
- **Supported Formats**: Text, Unicode text, Rich Text Format (RTF), HTML, images (bitmap, DIB, DIBV5), files, and more
- **Format Preservation**: All clipboard formats are preserved when pasting (bold, colors, sizes, etc.)
- **Memory Limits**: 
  - Max 5 formats per item
  - Max 5MB per individual format
  - Max 10MB total per clipboard item
  - Aggressive cleanup when exceeding 50 items
- **Hotkey**: Ctrl+NumPadDot (low-level keyboard hook for reliable detection)
- **Update Frequency**: Asynchronous clipboard monitoring with retry logic
- **Duplicate Detection**: Prevents adding identical consecutive items by comparing all formats and data
- **Sound**: Embedded click.mp3 plays on copy (fallback to file in exe directory); MCI warm-up ensures first copy plays sound
- **Navigation**: Arrow keys, Page Up/Down, Home/End for list navigation
- **Auto-hide**: List automatically hides when focus moves to another window
- **Clipboard persistence**: Saved to `%APPDATA%\clip2\history.dat` (binary format; up to 50 text items; load on startup, save on exit)
- **Snippets storage**: Registry `HKEY_CURRENT_USER\Software\clip2\Snippets`
- **Search**: Full-text search up to 500KB per item; finds matches across line breaks

## Requirements

- Windows API libraries: `user32`, `gdi32`, `shell32`, `winmm`, `comctl32`, `shlwapi`, `ole32`
- C++17 standard
- ico2.ico icon file (included)
- click.mp3 sound file (embedded as resource, with file fallback)

## Troubleshooting

- **Icon not showing**: Ensure `ico2.ico` is in the same directory or compiled as a resource
- **Hotkey not working**: Ensure NumLock is enabled or try the extended key combination
- **Sound not playing**: Ensure `click.mp3` is in the same directory as `clip2.exe` (or embedded as resource)
- **Performance issues**: The application uses minimal resources; if issues occur, check system resources

## Changelog

### Recent Features (Latest)
- ✅ **Clipboard persistence**: History saved to `%APPDATA%\clip2\history.dat`; survives app restarts (up to 50 text items)
- ✅ **Snippets overlay**: Press **SS** (S twice) to switch between clipboard and snippets; single-click to paste snippets
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
- Automatic cleanup when history exceeds 50 items
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

