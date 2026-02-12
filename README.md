# WindowMover

PowerShell utility to save and restore window layouts, sizes, and positions on Windows. Perfect for multi-monitor setups and quickly arranging your workspace.

## Overview

WindowMover consists of two complementary scripts:

| Script | Purpose |
|--------|---------|
| `winlayout_record.ps1` | Captures current window positions and saves to a config file |
| `winlayout_apply.ps1` | Restores windows to their saved positions from a config file |

## Quick Start

### 1. Record Your Layout
```powershell
# Record all visible windows
.\winlayout_record.ps1

# Record specific applications
.\winlayout_record.ps1 -ProcessName "chrome","code","notepad"

# Save to a specific file
.\winlayout_record.ps1 -Path "my-layout.json"
```

### 2. Apply Your Layout
```powershell
# Apply the default layout
.\winlayout_apply.ps1

# Apply a specific layout file
.\winlayout_apply.ps1 -Path "my-layout.json"
```

## winlayout_record.ps1

Records current window positions and saves them to a JSON configuration file.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Path` | string | `$env:USERPROFILE\windowlayout.config` | Output JSON file path |
| `CsvPath` | string | (none) | Optional CSV output path |
| `ProcessName` | string[] | (all) | Only record these processes |
| `ExcludeProcessName` | string[] | (none) | Exclude these processes |
| `Append` | switch | false | Append to existing file |
| `IncludeMinimized` | switch | false | Include minimized windows |
| `Deduplicate` | switch | false | Keep only largest window per group |
| `DedupBy` | string | `process` | Grouping: `process`, `process+title`, `monitor`, `process+monitor`, `process+title+monitor` |
| `DedupMonitorBy` | string | `device` | `device` or `index` |

### Examples

**Record all windows:**
```powershell
.\winlayout_record.ps1
```

**Record only VS Code and Chrome:**
```powershell
.\winlayout_record.ps1 -ProcessName "code","chrome"
```

**Record one window per process per monitor (multi-monitor):**
```powershell
.\winlayout_record.ps1 -Deduplicate -DedupBy "process+monitor"
```

**Append to existing config:**
```powershell
.\winlayout_record.ps1 -Append -Path "shared-layout.json"
```

**Export to CSV as well:**
```powershell
.\winlayout_record.ps1 -CsvPath "windows.csv"
```

## winlayout_apply.ps1

Applies saved window layouts to restore positions and sizes.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Path` | string | `$env:USERPROFILE\windowlayout.config` | Config file path |
| `BundleToApply` | string | (none) | Apply specific named bundle |
| `DryRun` | switch | false | Show what would be done without moving |
| `Record` | switch | false | Record mode (instead of apply) |
| `Process` | string[] | (none) | Processes to record (with `-Record`) |
| `RecordBundleName` | string | (none) | Save as named bundle (with `-Record`) |

### Examples

**Apply default layout:**
```powershell
.\winlayout_apply.ps1
```

**Apply specific file:**
```powershell
.\winlayout_apply.ps1 -Path "meeting-setup.json"
```

**Apply named bundle:**
```powershell
.\winlayout_apply.ps1 -BundleToApply "coding-setup"
```

**Test without moving (dry run):**
```powershell
.\winlayout_apply.ps1 -DryRun
```

**Record current layout as a bundle:**
```powershell
.\winlayout_apply.ps1 -Record -Process "chrome","code" -RecordBundleName "dev"
```

## Configuration File Format

### Simple Array Format
```json
[
  {
    "processName": "chrome",
    "title": "Google Chrome",
    "x": 0,
    "y": 0,
    "width": 1920,
    "height": 1080,
    "monitorIndex": 1
  },
  {
    "processName": "code",
    "title": "Visual Studio Code",
    "x": 1920,
    "y": 0,
    "width": 1920,
    "height": 1080,
    "monitorIndex": 2
  }
]
```

### Advanced Format with Bundles
```json
{
  "bundleDefaults": {
    "dpiMode": "auto",
    "pad": 10
  },
  "bundles": {
    "coding": [
      {
        "processName": "code",
        "preset": "LeftHalf",
        "monitorIndex": 1
      },
      {
        "processName": "chrome",
        "preset": "RightHalf", 
        "monitorIndex": 1
      }
    ],
    "meeting": [
      {
        "processName": "teams",
        "preset": "Full",
        "monitorIndex": 2
      }
    ]
  },
  "applyBundles": ["coding"]
}
```

### Entry Properties

| Property | Type | Description |
|----------|------|-------------|
| `processName` | string | Process name (required) |
| `x`, `y` | int | Position in pixels |
| `xPct`, `yPct` | float | Position as percentage |
| `width`, `height` | int | Size in pixels |
| `widthPct`, `heightPct` | float | Size as percentage |
| `monitorIndex` | int | Monitor number (0-based, -1 = primary) |
| `anchor` | string | Position anchor: `TopLeft`, `Top`, `TopRight`, `Left`, `Center`, `Right`, `BottomLeft`, `Bottom`, `BottomRight` |
| `preset` | string | Named preset (see below) |
| `grid` | string | Grid layout: `2x2`, `3x2`, etc. |
| `cell` | string | Grid cell: `row,col` |
| `pad` | int | Padding in pixels |
| `dpiMode` | string | `auto`, `logical`, or `physical` |
| `ensureRunning` | bool | Launch process if not running |
| `launchPath` | string | Path to executable for launching |
| `launchArgs` | string | Arguments for launching |
| `waitForSeconds` | int | Delay before targeting |
| `retryCount` | int | Retries to find window |
| `retryDelaySeconds` | int | Delay between retries |

## Presets

Use presets for common window arrangements:

| Preset | Description |
|--------|-------------|
| `Full` | Full screen |
| `LeftHalf` | Left 50% |
| `RightHalf` | Right 50% |
| `TopHalf` | Top 50% |
| `BottomHalf` | Bottom 50% |
| `LeftThird` | Left 33% |
| `CenterThird` | Center 33% |
| `RightThird` | Right 33% |
| `LeftTwoThirds` | Left 66% |
| `RightTwoThirds` | Right 66% |
| `TopLeftQuarter` | Top-left 25% |
| `TopRightQuarter` | Top-right 25% |
| `BottomLeftQuarter` | Bottom-left 25% |
| `BottomRightQuarter` | Bottom-right 25% |
| `CenteredLarge` | Centered 70% |

### Using Presets
```json
{
  "processName": "code",
  "preset": "LeftHalf",
  "monitorIndex": 0
}
```

## Grid Layouts

Define custom grid layouts:

```json
{
  "processName": "chrome",
  "grid": "3x2",
  "cell": "1,1",
  "monitorIndex": 0,
  "gutter": 10
}
```

Properties:
- `grid`: Format `rowsxcols` (e.g., `2x2`, `3x2`)
- `cell`: Format `row,col` (1-based)
- `rowSpan`: Rows to span (default 1)
- `colSpan`: Columns to span (default 1)
- `gutter`: Gap between cells in pixels
- `outerGutter`: Padding around grid

## Auto-Launch Applications

Automatically start applications if not running:

```json
{
  "processName": "outlook",
  "preset": "LeftHalf",
  "ensureRunning": true,
  "launchPath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
  "launchArgs": "",
  "postLaunchDelaySeconds": 5,
  "launchTimeoutSeconds": 30
}
```

## DPI Scaling

Handle different display DPIs:

| Mode | Description |
|------|-------------|
| `auto` | Scale based on target window's DPI (default) |
| `logical` | No scaling |
| `physical` | Always scale using system DPI |

```json
{
  "processName": "chrome",
  "preset": "LeftHalf",
  "dpiMode": "auto"
}
```

## Requirements

- Windows 8 or later
- PowerShell 5.1 or PowerShell 7+
- .NET Framework 4.5+ (for WinForms)

## Tips

### Create Desktop Shortcuts
Create shortcuts to quickly apply layouts:
```
PowerShell.exe -File "C:\Tools\WindowMover\winlayout_apply.ps1" -BundleToApply "coding"
```

### Scheduled Tasks
Run at login to restore your preferred layout:
```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Tools\WindowMover\winlayout_apply.ps1"
$trigger = New-ScheduledTaskTrigger -AtLogon
Register-ScheduledTask -TaskName "Restore Window Layout" -Action $action -Trigger $trigger
```

### Multiple Layouts
Create different configs for different scenarios:
- `coding.json` - IDE and browser side by side
- `meeting.json` - Teams/Zoom full screen on second monitor
- `admin.json` - Multiple PowerShell windows

## Troubleshooting

### "No matching windows found"
- Ensure the application has a visible window (not just a tray icon)
- Try with `-IncludeMinimized` flag
- Check the process name with `Get-Process`

### Windows not moving correctly
- Check DPI settings with `-Verbose`
- Try different `dpiMode` values
- Verify monitor indices match your setup

### Application doesn't launch
- Verify `launchPath` points to the correct executable
- Check for typos in `processName`
- Ensure the path has proper escaping in JSON

### Layout applies to wrong monitor
- Monitor indices are 0-based and may change with display arrangement
- Use `monitorIndex` in config to specify target

## License

MIT License - Free to use and modify.
