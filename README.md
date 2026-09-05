# clip2

A fast Windows clipboard manager that lives in the system tray. Open a dual-pane AMOLED overlay to search, pin, transform, and paste history — with snippets, smart paste modes, and encrypted on-disk persistence.

## Requirements

- Windows 10 or 11
- CMake 3.15+
- A C++17 toolchain (Visual Studio 2019/2022, or MinGW)

## Build

Easiest:

```bat
build.bat
```

That configures into `build/`, builds Release, and copies `clip2.exe` to the repo root.

Or with CMake directly:

```bat
   mkdir build
   cd build
   cmake .. -G "Visual Studio 17 2022" -A x64
   cmake --build . --config Release
   ```

Linked libraries include `user32`, `gdi32`, `shell32`, `winmm`, `comctl32`, `shlwapi`, `ole32`, `oleaut32`, `comdlg32`, `crypt32`, and Media Foundation (`mfplat`, `mf`, `mfreadwrite`, `mfuuid`) for the click sound. If the Windows SDK’s UIAutomation library is found, the build defines `HAVE_UIAUTOMATION` and enables **Copy from focused control**.

## Run

Start `clip2.exe`. It appears in the system tray (icon from `ico2.ico`). Left-click the tray icon to show or hide the clipboard overlay. Right-click for the tray menu.

## Default hotkeys

| Action | Default | Configurable |
|--------|---------|--------------|
| Toggle clipboard overlay | **Ctrl+NumPadDot** | Settings |
| Toggle snippets overlay | *(none)* | Settings |
| Copy from focused control | **Ctrl+F10** | Settings |
| Paste as keystrokes (no clipboard swap) | **Ctrl+F11** | Settings |
| Paste via clipboard + Ctrl+V | **Ctrl+Shift+F11** | Settings |

All five can be rebound under **Settings**. Defaults restores the table above (snippets stays unbound).

## Overlay

Two panels:

- **Main** — full clipboard history (or snippets when switched)
- **Pinned** (left) — pinned items only

**Tab** moves focus between panels. Drag either panel; the pair stays linked and the position is remembered across sessions.

### Clipboard mode (main panel)

| Input | Action |
|-------|--------|
| **Enter** | Paste selection (or multi-selection) |
| Digits **0–9** | Type an item number; that item jumps to the top so you can preview/act on it; **Enter** pastes it |
| **Ctrl+Click** / **Shift+Arrows** | Multi-select |
| **Ctrl** while pasting | Plain text only |
| **Delete** | Remove the selected item |
| **Ctrl+F** | Focus search |
| **Ctrl+Right** / **Ctrl+Left** | Switch to snippets / back to clipboard |
| **Esc** | Close overlay |
| Double-click | Paste item |

Search uses fuzzy matching (subsequence + ranking) over full item text (up to 500 KB), with a per-item trigram index so large histories stay responsive.

Hover an image/video for a larger preview. Row thumbnails load lazily when a row becomes visible; the inactive panel draws type badges instead of decoding images.

### Smart paste (list focused, not typing in search)

| Key | Action |
|-----|--------|
| **U** | Paste URL with tracking params stripped (`utm_*`, `fbclid`, `gclid`, …) |
| **M** | **If 2+ items are multi-selected:** merge them into a **new top history item** (Unicode + RTF + HTML, unpinned). **Otherwise:** paste as a Markdown link |
| **P** | Paste as **plain text**. **If 2+ items are multi-selected:** also insert that plain merge as a new top history item (one line per item). File-path paste remains in the context menu. |
| **H** | Paste HTML as plain text |
| **E** | Edit selection (or joined multi-selection), then paste — history unchanged |
| **X** | Same editor; **Save** / **Ctrl+Enter** adds a **new** history item and updates the clipboard |
| **Z** | **Excel cell fill (needs 2+ multi-selected items):** for each item, **F2** (edit cell) → paste rich formats → **Enter** (commit and move down). Same cycle for every selected item. |

Also available from the right-click menu (**Paste as…**, **Edit & Paste…**, **Edit & Save as new…**, **Merge selection into new item**, **Paste & merge as plain text**, **Paste into Excel**). Modes that don’t apply fail silently.

### Multi-paste & merge

- Multi-select text/RTF/HTML items and press **Enter**: one clipboard payload with each item on its own line (`\r\n` / RTF paragraph / HTML `<br>`), formatting preserved via RichEdit merge, single Ctrl+V (no Enter keystrokes). The same merged payload is also inserted as a **new top history item** (rich, or plain when using Ctrl+Enter).
- Mixed image/file selections paste sequentially, with line breaks inserted as clipboard text where needed; joined text (when present) is recorded as a new plain history item.
- **M** with 2+ selected items creates a new unpinned history entry at the top of the main list (rich formats kept).
- **P** pastes as plain text (single item or multi-select). With 2+ selected items it also inserts a plain merged item at the top.
- **Z** with 2+ selected items is the Excel special paste: for each item, **F2** opens the cell for edit, the item is pasted with its original rich clipboard formats, then **Enter** commits and moves down one row. Same F2 → paste → Enter cycle for every selected item. Independent of Enter multi-paste — nothing is merged and no line breaks are embedded.

### Pinned items

Right-click → **Pin** / **Unpin**. Pinned items appear in the left panel; in the main list they keep their normal history order. They survive **Clear List** and are always saved. Shown with a left accent bar.

### Context menu (clipboard items)

Pin/Unpin, Copy as plain text, Open URL (http/https), Save image to file…, Keystroke paste, Transform Text (upper/lower/title case, remove line breaks, trim, plain text), Paste as…, Edit…, Delete, Clear List.

## Snippets mode

Open with **Ctrl+Right** from the clipboard overlay, the optional snippets hotkey, or the tray **Snippets** menu.

| Input | Action |
|-------|--------|
| **Enter** / single-click | Paste snippet |
| **A** | Add snippet (popup; list focused, not search) |
| **E** | Edit selected snippet (popup) |
| Type | Filters via search |
| `*set` + **Enter** | Open Manage Snippets |
| **Ctrl+Left** | Back to clipboard |
| **Esc** | Close |

Snippets support RTF (Rich Edit) and placeholders:

`{{date}}` `{{time}}` `{{datetime}}` `{{year}}` `{{month}}` `{{day}}` `{{hour}}` `{{minute}}` `{{second}}` `{{clipboard}}`

`*set` cannot be used as a snippet name. Snippets are stored under `HKEY_CURRENT_USER\Software\clip2\Snippets`.

## Tray menu

- **Show Clipboard**
- **Copy from focused control** — only if built with UIAutomation
- **Snippets** — paste a snippet, or open the manager
- **Start with Windows** — toggle Run key `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\clip2`
- **Settings** — hotkeys, theme, fonts, colors, history size
- **Restart**
- **Exit**

## Settings

**Shortcuts & Theme** dialog:

- Rebind the five hotkeys (click a field, press the combo). **Defaults** restores built-in shortcuts.
- Overlay theme (AMOLED neon): Neon Green (default), Red, Blue, Cyan, Purple, Yellow, Orange, White.
- Overlay font face (all installed fonts; default Consolas) and optional font color override.
- History size: **10–2000** items (default **300**).
- Per-element colors: Background, Text, Accent, Selected text, Border, Dim — click to pick, double-click to reset one slot; **Defaults** clears all color overrides.

Theme, font, and colors apply live. Hotkeys and history size apply on **Save**.

## Persistence

| What | Where |
|------|--------|
| History | `%APPDATA%\clip2\history.dat` |
| Large formats (≥ 256 KB) | `%APPDATA%\clip2\blobs\<sha1>` |
| Settings / hotkeys / theme / position | `HKEY_CURRENT_USER\Software\clip2` |
| Snippets | `HKEY_CURRENT_USER\Software\clip2\Snippets` |

History file format:

- Outer container **CLP3**: DPAPI-encrypted per Windows user.
- Inner payload **CLP2** v2: pinned flag + all formats per item.
- Legacy plaintext CLP2 loads and upgrades on the next save.

Saves are debounced (~1.5 s after the last change), written atomically (`history.dat.tmp` then rename), and flushed on exit. Unreferenced blob files are removed after a successful save. Pinned items are always persisted; other items are capped by the configured history size.

Capture limits (runtime): up to 12 formats per item, under 5 MB per format, 10 MB total per item. Search indexes up to 500 KB of text per item.

## Copy from focused control

When an app never puts text on the clipboard:

1. Tray → **Copy from focused control**, or **Ctrl+F10**
2. clip2 tries UI Automation (if built with it), then a synthetic Ctrl+C / Ctrl+A+Ctrl+C fallback
3. Captured text is added to history (and the clipboard when the UIA path succeeds)

If UIAutomation was not found at configure time, the tray item is omitted; rebuild with the Windows SDK to enable it.

## Paste without / with clipboard swap

- **Ctrl+F11** — types the selected (or newest) history item as Unicode keystrokes into the focused app (does not replace the system clipboard).
- **Ctrl+Shift+F11** — puts the item’s formats on the clipboard and sends Ctrl+V.

## Developer

- **Developer:** Unnamed10110
- **Email:** trojan.v6@gmail.com
- **Alternative:** sergiobritos10110@gmail.com

## License

Provided as-is for personal use.
