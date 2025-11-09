# clip2

A fast and modern clipboard manager for Windows with OLED theme and rich format support.

## Features

- **Background Operation**: Runs silently in the system tray
- **Rich Format Support**: Preserves formatting (bold, colors, sizes, etc.) when copying and pasting
- **Multiple Formats**: Supports text, rich text, images, files, and more
- **History Management**: Stores up to 1000 clipboard items
- **Modern UI**: Black OLED theme with bright red accents and rounded borders
- **Quick Access**: Press `Ctrl+NumPadDot` to show/hide the clipboard list
- **Search**: Real-time search filtering with `Ctrl+F`
- **Keyboard Navigation**: Arrow keys to navigate, Enter to paste
- **Image Preview**: Hover over image/video items to see larger previews
- **Duplicate Prevention**: Automatically prevents duplicate consecutive copies
- **Start with Windows**: Option to start automatically with Windows session

## Usage

### Basic Operations

- **Show/Hide List**: Press `Ctrl+NumPadDot`
- **Paste Item**: 
  - Type a number and press `Enter` (e.g., type `5` then `Enter` to paste item 5)
  - Press a single digit (1-9) for quick paste
  - Double-click an item
  - Use arrow keys to select and press `Enter`
- **Search**: Press `Ctrl+F` to focus search box, type to filter, press `Enter` to paste first result
- **Plain Text Paste**: Hold `Ctrl` while pasting to paste as plain text only
- **Clear List**: Right-click the list and select "Clear List"

### Tray Icon Menu

Right-click the tray icon to access:
- **Show Clipboard**: Show the clipboard list
- **Start with Windows**: Toggle automatic startup (checkmark indicates enabled)
- **Exit**: Close the application

## Requirements

- Windows 10 or later
- Visual C++ Redistributable (if not included)

## Building

### Prerequisites

- CMake 3.15 or later
- C++17 compatible compiler (MSVC, MinGW, etc.)

### Build Instructions

```bash
mkdir build
cd build
cmake ..
cmake --build . --config Release
```

The executable will be created as `clip2.exe` in the `build` directory.

## Files

- `ClipboardManager.cpp` / `ClipboardManager.h`: Main application code
- `CMakeLists.txt`: Build configuration
- `ClipboardManager.rc`: Resource file for icon
- `misc02.ico`: Application icon
- `click.mp3`: Sound effect played when copying

## Developer

**Unnamed10110**
- Email: trojan.v6@gmail.com
- Email: sergiobritos10110@gmail.com

## License

This project is provided as-is for personal use.
