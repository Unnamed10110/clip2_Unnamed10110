# clip2

A modern clipboard manager for Windows with OLED theme and advanced features.


## Features

- ✅ **Background Operation**: Runs silently in the system tray
- ✅ **Universal Clipboard Support**: Captures text, rich text, images, files, and more
- ✅ **LIFO History**: Most recently copied item appears first (up to 1000 items)
- ✅ **Hotkey Access**: Press **Ctrl+NumPadDot** to show/hide the clipboard list
- ✅ **Quick Paste**: Press **1-9** or type number + Enter to paste items from the list
- ✅ **Visual Feedback**: Click sound plays when items are copied
- ✅ **System Tray Icon**: Always visible when running (uses misc02.ico)
- ✅ **Modern OLED Theme**: Black background with bright red accents and rounded borders
- ✅ **Search Functionality**: Filter clipboard history with Ctrl+F
- ✅ **Image/Video Thumbnails**: Visual previews for images and videos
- ✅ **Format Preservation**: Pastes with original formatting (bold, colors, etc.)
- ✅ **Plain Text Mode**: Hold Ctrl while pasting for plain text
- ✅ **Start with Windows**: Option to auto-start with Windows session
- ✅ **Duplicate Prevention**: Automatically skips identical consecutive copies
- ✅ **Stability**: Robust error handling prevents crashes with large or complex clipboard data
- ✅ **Fast Performance**: Optimized for minimal CPU usage and memory footprint

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

5. **Hide the list**: Press **Ctrl+NumPadDot** again, **Escape**, or click outside the window

6. **System tray menu**: Right-click the tray icon for options
   - **Show Clipboard**: Show the clipboard list
   - **Start with Windows**: Toggle auto-start with Windows (checkmark indicates enabled)
   - **Exit**: Close the application

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
- **Sound**: Embedded click.mp3 plays on copy (with fallback to file)
- **Navigation**: Arrow keys, Page Up/Down, Home/End for list navigation
- **Auto-hide**: List automatically hides when focus moves to another window

## Requirements

- Windows API libraries: `user32`, `gdi32`, `shell32`, `winmm`, `comctl32`, `shlwapi`, `ole32`
- C++17 standard
- misc02.ico icon file (included)
- click.mp3 sound file (embedded as resource, with file fallback)

## Troubleshooting

- **Icon not showing**: Ensure `misc02.ico` is in the same directory or compiled as a resource
- **Hotkey not working**: Ensure NumLock is enabled or try the extended key combination
- **Sound not playing**: Ensure `click.mp3` is in the same directory as `clip2.exe` (or embedded as resource)
- **Performance issues**: The application uses minimal resources; if issues occur, check system resources

## Recent Improvements & Fixes

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

