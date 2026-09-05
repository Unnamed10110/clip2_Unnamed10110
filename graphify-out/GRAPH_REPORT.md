# Graph Report - clip2_Unnamed10110  (2026-09-05)

## Corpus Check
- Corpus is ~47,012 words - fits in a single context window. You may not need a graph.

## Summary
- 473 nodes · 1379 edges · 19 communities
- Extraction: 81% EXTRACTED · 19% INFERRED · 0% AMBIGUOUS · INFERRED: 266 edges (avg confidence: 0.82)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Manager State Fields
- Rich Text Helpers
- Dialogs and Hotkeys
- Overlay List Actions
- Build Docs Overlay
- ClipboardItem Data Model
- Synthetic Input Paste
- History Persistence
- UIA Text Extraction
- Startup and Tray
- Theme Colors
- Lifecycle and Instance
- Clipboard Ingest Dedup
- Search Index
- Image Preview Bitmaps
- Pane State Types
- RichEdit Callbacks
- Overlay Back Buffer

## God Nodes (most connected - your core abstractions)
1. `ClipboardManager` - 324 edges
2. `size` - 85 edges
3. `ClipboardItem` - 75 edges
4. `ListWindowProc` - 56 edges
5. `data` - 40 edges
6. `ThemeGdi` - 28 edges
7. `WindowProc` - 26 edges
8. `GetPlainTextForDirectPaste()` - 21 edges
9. `MergeClipboardItemsRich()` - 21 edges
10. `FilterItems` - 21 edges

## Surprising Connections (you probably didn't know these)
- `Pencil Edit Icon` --conceptually_related_to--> `Edit and Paste`  [INFERRED]
  misc02.png → README.md
- `SmartPasteItem` --references--> `ThemePreset`  [INFERRED]
  ClipboardManager.h → ClipboardManager.cpp
- `ListWindowProc` --references--> `ThemeGdi`  [INFERRED]
  ClipboardManager.h → ClipboardManager.cpp
- `MergeSelectedItems` --references--> `ThemeGdi`  [INFERRED]
  ClipboardManager.h → ClipboardManager.cpp
- `PasteItem` --references--> `ThemeGdi`  [INFERRED]
  ClipboardManager.h → ClipboardManager.cpp

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Encrypted History Persistence Pipeline** — readme_clp3, readme_clp2, readme_dpapi, readme_blob_storage, readme_atomic_save [EXTRACTED 1.00]
- **Copy from Focused Control Feature** — cmakelists_uiautomation, cmakelists_have_uiautomation, readme_copy_from_focused_control [EXTRACTED 1.00]
- **Clipboard Overlay Interaction Surface** — readme_dual_pane_overlay, readme_pinned_items, readme_fuzzy_search, readme_smart_paste [EXTRACTED 1.00]

## Communities (19 total, 0 thin omitted)

### Community 0 - "Manager State Fields"
Cohesion: 0.02
Nodes (87): ClipboardManager, activeIsPinned, clipboardHistory, copyFocusedHotkey, DEFAULT_MAX_ITEMS, editPasteSaveAsNew, filteredIndices, filteredIndicesBase (+79 more)

### Community 1 - "Rich Text Helpers"
Cohesion: 0.11
Nodes (58): AppendDecodedHtmlEntity(), BuildHtmlClipboardFormat(), BuildRtfWrapFromUnicode(), CfHtml(), CfRtf(), HydrateAllFormats, CopyItemAsPlainText, ExpandSnippetPlaceholders (+50 more)

### Community 2 - "Dialogs and Hotkeys"
Cohesion: 0.14
Nodes (34): EditPasteDialogProc, EditPasteEditProc, HotkeyFromKeyMessage, HotkeyToString, IsOverlayWindow, IsStartupWithWindows, LowLevelKeyboardProc, SaveHotkeyConfig (+26 more)

### Community 3 - "Overlay List Actions"
Cohesion: 0.13
Nodes (33): AcquireBackBufferDC(), ActiveListHwnd, ApplyNumberInputPromotion, ClampOverlayPosition, ClearClipboardHistory, ClearMultiSelection, CommitCapturedText, DeleteItem (+25 more)

### Community 4 - "Build Docs Overlay"
Cohesion: 0.08
Nodes (32): Post-Build click.mp3 Copy, CMake Project clip2, HAVE_UIAUTOMATION Compile Definition, RC-Embedded App Manifest, Media Foundation Click Sound Decode, Optional UIAutomation Link, WIN32 GUI Executable Target, Pencil Edit Icon (+24 more)

### Community 5 - "ClipboardItem Data Model"
Cohesion: 0.07
Nodes (28): ClipboardItem, blobOrigin, charBloom, DehydrateBlobFormats, fileType, format, formatName, formats (+20 more)

### Community 6 - "Synthetic Input Paste"
Cohesion: 0.07
Nodes (29): CopyFocusedViaSyntheticCopy, ReadClipboardUnicodeText(), ReleaseAllModifierKeysForKeystrokePaste(), SendCtrlA(), SendCtrlC(), SendCtrlKey(), SendEnterKey(), SendF2Key() (+21 more)

### Community 7 - "History Persistence"
Cohesion: 0.16
Nodes (28): AppendBytes(), AppendDword(), BackupClipboardSerialFormats(), GetPreview, HydrateFormat, LoadClipboardHistory, SaveClipboardHistory, TryCaptureClipboardImmediately (+20 more)

### Community 8 - "UIA Text Extraction"
Cohesion: 0.15
Nodes (27): BSTR, CharBloomBit(), ComputeFilteredForPane, CopyFromFocusedControlViaUIA, WarmUpClickSound, time_point, wstring, ExtractElementText() (+19 more)

### Community 9 - "Startup and Tray"
Cohesion: 0.11
Nodes (21): CreateTrayIcon, Initialize, InstallKeyboardHook, LoadHotkeyConfig, LoadSnippets, RegisterHotkey, SetThemeFontFace, EnsureOverlayFontVariants() (+13 more)

### Community 10 - "Theme Colors"
Cohesion: 0.14
Nodes (20): ApplyThemeId(), BlendColor(), RefreshThemeVisuals, ResetThemeColorOverrides, SetTheme, SetThemeColorOverride, SetThemeFontColor, EffectiveThemeColor() (+12 more)

### Community 11 - "Lifecycle and Instance"
Cohesion: 0.15
Nodes (13): AcquireSingleInstanceLock(), ClipboardManager::ClipboardManager(), RemoveTrayIcon, Run, Stop, UninstallKeyboardHook, UnregisterHotkey, ReleaseSingleInstanceLock() (+5 more)

### Community 12 - "Clipboard Ingest Dedup"
Cohesion: 0.24
Nodes (10): AreDuplicateClipboardItems(), AreDuplicateImages(), CfPng(), GetFileType, ProcessClipboard, ProcessClipboardFromSnapshot, SaveImageItemToFile, TrimHistory (+2 more)

### Community 13 - "Search Index"
Cohesion: 0.24
Nodes (5): GetFullSearchableText, RebuildSearchIndex, DeferredPrimary, BYTE, UINT

### Community 14 - "Image Preview Bitmaps"
Cohesion: 0.24
Nodes (8): CreateBitmapFromData, EnsureThumbnail, GeneratePreviewBitmap, GenerateThumbnail, GetPreviewBitmap, ShowPreviewWindow, TrimBitmapCaches, HBITMAP

### Community 15 - "Pane State Types"
Cohesion: 0.20
Nodes (9): map, vector, wstring, PaneState, filteredIndices, scrollOffset, searchText, selectedIndex (+1 more)

### Community 16 - "RichEdit Callbacks"
Cohesion: 0.36
Nodes (8): Clip2ExceptionHandler(), DWORD, RichEditStreamInCallback(), RichEditStreamOutCallback(), DWORD_PTR, LONG, LPBYTE, LPEXCEPTION_POINTERS

### Community 17 - "Overlay Back Buffer"
Cohesion: 0.29
Nodes (7): OverlayBackBuffer, bmp, dc, h, oldBmp, w, HGDIOBJ

## Knowledge Gaps
- **144 isolated node(s):** `name`, `txt`, `selBg`, `border`, `dim` (+139 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 174 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `ClipboardManager` connect `Manager State Fields` to `Rich Text Helpers`, `Dialogs and Hotkeys`, `Overlay List Actions`, `ClipboardItem Data Model`, `Synthetic Input Paste`, `History Persistence`, `UIA Text Extraction`, `Startup and Tray`, `Theme Colors`, `Lifecycle and Instance`, `Clipboard Ingest Dedup`, `Search Index`, `Image Preview Bitmaps`, `Pane State Types`, `RichEdit Callbacks`, `Overlay Back Buffer`?**
  _High betweenness centrality (0.734) - this node is a cross-community bridge._
- **Why does `ClipboardItem` connect `ClipboardItem Data Model` to `Manager State Fields`, `Rich Text Helpers`, `History Persistence`, `UIA Text Extraction`, `Clipboard Ingest Dedup`, `Search Index`, `Image Preview Bitmaps`, `Pane State Types`?**
  _High betweenness centrality (0.122) - this node is a cross-community bridge._
- **Why does `ThemeGdi` connect `Synthetic Input Paste` to `Manager State Fields`, `Rich Text Helpers`, `Overlay List Actions`, `UIA Text Extraction`, `Theme Colors`?**
  _High betweenness centrality (0.063) - this node is a cross-community bridge._
- **Are the 4 inferred relationships involving `ClipboardItem` (e.g. with `ComputeFilteredForPane` and `PasteItem`) actually correct?**
  _`ClipboardItem` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 41 inferred relationships involving `ListWindowProc` (e.g. with `EnsureThumbnail` and `GetPreviewBitmap`) actually correct?**
  _`ListWindowProc` has 41 INFERRED edges - model-reasoned connections that need verification._
- **What connects `name`, `txt`, `selBg` to the rest of the system?**
  _144 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Manager State Fields` be split into smaller, more focused modules?**
  _Cohesion score 0.022988505747126436 - nodes in this community are weakly interconnected._