# Auto-Editor GUI

A Windows GUI for [auto-editor](https://github.com/wyattblue/auto-editor), a command-line tool that automatically removes silent or inactive sections from video files. This script wraps auto-editor with a drag-and-drop interface and sliders for the most common settings.

## Features

- Adjust audio threshold, margin before/after, and target video editor
- Drag and drop video files to process them instantly
- Saves your settings between sessions in `auto-editor-config.json`
- One-click update when a newer version of auto-editor is available on PyPI

---

## Prerequisites

- **Windows 10 or 11**
- **Python 3.8+** — [download from python.org](https://www.python.org/downloads/)  
  During installation, check **"Add Python to PATH"**
- **auto-editor** — install via pip after Python is set up:
  ```
  pip install auto-editor
  ```

---

## Installing to the Windows Start Menu

To make Auto-Editor searchable from the Windows Start bar like a regular app, run the PowerShell command below. 

> **Why not use the shortcut Properties dialog?** Windows parses the Target field and separates the executable from its arguments. When the `-File` path contains spaces, it mangles the result even with quotes. Creating the shortcut via PowerShell sets those fields separately and avoids the issue entirely.

Open a PowerShell window **in the same folder as this README** (right-click the folder in Explorer → **Open in Terminal**), then paste and run:

```powershell
$scriptDir = (Get-Location).Path
$shell     = New-Object -ComObject WScript.Shell
$shortcut  = $shell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Auto-Editor.lnk")
$shortcut.TargetPath       = "powershell.exe"
$shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptDir\AutoEditorGui.ps1`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.IconLocation     = "$scriptDir\AutoEditorScript.ico"
$shortcut.Save()
```

Auto-Editor will appear in the Windows Start bar as soon as the script completes.

---

## Settings

| Control | Description |
|---|---|
| Audio Threshold | How loud audio must be to be considered "active" (percentage). Higher = only loud audio is kept. |
| Margin Before | Seconds of audio to keep before each active section. |
| Margin After | Seconds of audio to keep after each active section. |
| Video Editor | Export format — `resolve` (DaVinci Resolve), `final-cut-pro`, or `shotcut`. |

Settings are saved to `auto-editor-config.json` in the same folder as the script. Delete this file to reset to defaults.

---

## Updating auto-editor

The title bar shows your installed version and the latest version available on PyPI. If they differ, the **Update** button becomes active. Click it to run `pip install auto-editor --upgrade` automatically.

---

## Files

| File | Purpose |
|---|---|
| `AutoEditorGui.ps1` | Main GUI script |
| `AutoEditorScript.ico` | Icon used by the Start Menu shortcut |
| `auto-editor-config.json` | Saved settings (created on first save) |
