<#
.SYNOPSIS
    Records current window positions and saves them to a JSON configuration file.

.DESCRIPTION
    Captures the position, size, and monitor of visible top-level windows.
    Supports filtering by process name, deduplication, and multiple output formats.
    Use this script to create layout configs that can be restored later with winlayout_apply.ps1.

.PARAMETER Path
    Destination JSON file path. Default: "$env:USERPROFILE\windowlayout.config"

.PARAMETER CsvPath
    Optional CSV output path written alongside JSON.

.PARAMETER ProcessName
    Only include specific process names (e.g., "chrome","notepad","Code").

.PARAMETER ExcludeProcessName
    Exclude specific process names.

.PARAMETER Append
    Append captured entries to existing JSON file (merges arrays).

.PARAMETER IncludeMinimized
    Include minimized windows in the capture.

.PARAMETER Deduplicate
    Keep only the largest (by area) window per group.

.PARAMETER DedupBy
    Deduplication grouping method:
    - process: largest per process
    - process+title: largest per process+title
    - monitor: largest per monitor
    - process+monitor: largest per process on each monitor
    - process+title+monitor: largest per process+title on each monitor
    Default: process

.PARAMETER DedupMonitorBy
    For per-monitor dedup, group by 'device' (default) or 'index'.

.EXAMPLE
    .\winlayout_record.ps1
    Records all visible windows to default location.

.EXAMPLE
    .\winlayout_record.ps1 -ProcessName "chrome","code" -Path "dev-layout.json"
    Records only Chrome and VS Code windows.

.EXAMPLE
    .\winlayout_record.ps1 -Deduplicate -DedupBy "process+monitor"
    Records one window per process per monitor (useful for multi-monitor setups).

.NOTES
    Version: 2.1
    Requires: Windows 8+ with PowerShell 5.1+
    
    Exit Codes:
    0 - Success
    1 - General error
#>

[CmdletBinding()]
param(
    [string]$Path = (Join-Path $env:USERPROFILE 'windowlayout.config'),
    [string]$CsvPath,
    [string[]]$ProcessName,
    [string[]]$ExcludeProcessName,
    [switch]$Append,
    [switch]$IncludeMinimized,
    [switch]$Deduplicate,
    [ValidateSet('process','process+title','monitor','process+monitor','process+title+monitor')]
    [string]$DedupBy = 'process',
    [ValidateSet('device','index')]
    [string]$DedupMonitorBy = 'device'
)

#region Win32 Interop
$null = @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class Win32Window {
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    public static List<IntPtr> EnumAllWindows() {
        var list = new List<IntPtr>();
        EnumWindows((hWnd, lParam) => { list.Add(hWnd); return true; }, IntPtr.Zero);
        return list;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left; public int Top; public int Right; public int Bottom;
}
'@ | Add-Type -PassThru -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Windows.Forms | Out-Null
#endregion

#region Helper Functions
function Get-MonitorInfo {
    param([IntPtr]$Handle)
    $screen = [System.Windows.Forms.Screen]::FromHandle($Handle)
    $index = [array]::IndexOf([System.Windows.Forms.Screen]::AllScreens, $screen) + 1
    return @{
        Device = $screen.DeviceName
        Index = $index
        Bounds = $screen.Bounds
        WorkingArea = $screen.WorkingArea
    }
}

function Get-WindowList {
    param([string[]]$Include, [string[]]$Exclude, [switch]$Minimized)
    
    $includeSet = if ($Include) { $Include | ForEach-Object { $_.ToLower() } } else { @() }
    $excludeSet = if ($Exclude) { $Exclude | ForEach-Object { $_.ToLower() } } else { @() }
    
    $results = [System.Collections.Generic.List[object]]::new()
    $handles = [Win32Window]::EnumAllWindows()
    
    foreach ($hWnd in $handles) {
        if (-not [Win32Window]::IsWindowVisible($hWnd)) { continue }
        
        $len = [Win32Window]::GetWindowTextLength($hWnd)
        if ($len -le 0) { continue }
        
        $titleBuilder = [System.Text.StringBuilder]::new($len + 1)
        [void][Win32Window]::GetWindowText($hWnd, $titleBuilder, $titleBuilder.Capacity)
        $title = $titleBuilder.ToString()
        
        if ([string]::IsNullOrWhiteSpace($title)) { continue }
        
        $processId = 0
        [void][Win32Window]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        if ($processId -eq 0) { continue }
        
        try { $proc = Get-Process -Id $processId -ErrorAction Stop } catch { continue }
        $procName = $proc.ProcessName
        
        if ($includeSet.Count -gt 0 -and $includeSet -notcontains $procName.ToLower()) { continue }
        if ($excludeSet.Count -gt 0 -and $excludeSet -contains $procName.ToLower()) { continue }
        
        $rect = New-Object RECT
        if (-not [Win32Window]::GetWindowRect($hWnd, [ref]$rect)) { continue }
        
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        if ($width -le 0 -or $height -le 0) { continue }
        
        if (-not $Minimized -and [Win32Window]::IsIconic($hWnd)) { continue }
        
        $mon = Get-MonitorInfo -Handle $hWnd
        
        $results.Add([PSCustomObject]@{
            ProcessName = $procName
            Title = $title
            X = [int]$rect.Left
            Y = [int]$rect.Top
            Width = [int]$width
            Height = [int]$height
            Area = [int]($width * $height)
            MonitorIndex = [int]$mon.Index
            MonitorDevice = [string]$mon.Device
        }) | Out-Null
    }
    
    return $results | Sort-Object MonitorIndex, ProcessName, Title
}

function Remove-Duplicates {
    param([System.Collections.IEnumerable]$Entries, [string]$By, [string]$MonitorKey)
    
    if (-not $Entries) { return @() }
    
    $monitorSelector = if ($MonitorKey -eq 'index') { 
        { $_.MonitorIndex } 
    } else { 
        { $_.MonitorDevice } 
    }
    
    switch ($By) {
        'process' { 
            $groups = $Entries | Group-Object ProcessName 
        }
        'process+title' { 
            $groups = $Entries | Group-Object { "{0}||{1}" -f $_.ProcessName, $_.Title }
        }
        'monitor' { 
            $groups = $Entries | Group-Object (& $monitorSelector.InvokeReturnAsIs($_))
        }
        'process+monitor' { 
            $groups = $Entries | Group-Object { 
                "{0}||{1}" -f $_.ProcessName, (& $monitorSelector.InvokeReturnAsIs($_))
            }
        }
        'process+title+monitor' { 
            $groups = $Entries | Group-Object { 
                "{0}||{1}||{2}" -f $_.ProcessName, $_.Title, (& $monitorSelector.InvokeReturnAsIs($_))
            }
        }
    }
    
    return $groups | ForEach-Object { 
        $_.Group | Sort-Object Area -Descending | Select-Object -First 1 
    }
}

function Export-CsvData {
    param([System.Collections.IEnumerable]$Entries, [string]$CsvPath)
    if (-not $CsvPath) { return }
    try {
        $Entries | Select-Object ProcessName, Title, X, Y, Width, Height, MonitorIndex, MonitorDevice |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV exported: $CsvPath" -ForegroundColor Green
    } catch {
        Write-Warning "CSV export failed: $($_.Exception.Message)"
    }
}
#endregion

#region Main
$ErrorActionPreference = 'Stop'

try {
    Write-Host "Recording window layout..." -ForegroundColor Cyan
    
    $entries = Get-WindowList -Include $ProcessName -Exclude $ExcludeProcessName -Minimized:$IncludeMinimized
    
    if (-not $entries -or $entries.Count -eq 0) {
        Write-Warning "No matching windows found."
        exit 0
    }
    
    if ($Deduplicate) {
        $beforeCount = $entries.Count
        $entries = Remove-Duplicates -Entries $entries -By $DedupBy -MonitorKey $DedupMonitorBy
        Write-Host "Deduplicated: $beforeCount -> $($entries.Count) entries ($DedupBy, monitor=$DedupMonitorBy)" -ForegroundColor Cyan
    }
    
    # Remove helper property before output
    $output = $entries | Select-Object ProcessName, Title, X, Y, Width, Height, MonitorIndex, MonitorDevice
    
    Export-CsvData -Entries $output -CsvPath $CsvPath
    
    if ($Append -and (Test-Path -LiteralPath $Path)) {
        try {
            $existing = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
            $merged = if ($existing -is [System.Collections.IEnumerable]) { 
                @($existing) + @($output) 
            } else { 
                @($existing) + @($output) 
            }
            $merged | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
            Write-Host "Appended $($output.Count) entries to $Path" -ForegroundColor Green
        } catch {
            Write-Warning "Append failed, creating new file: $($_.Exception.Message)"
            $output | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
            Write-Host "Created new file: $Path" -ForegroundColor Green
        }
    } else {
        if (Test-Path -LiteralPath $Path) {
            $backup = "$Path.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            Copy-Item -LiteralPath $Path -Destination $backup -ErrorAction SilentlyContinue
            Write-Host "Backup created: $backup" -ForegroundColor DarkGray
        }
        $output | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
        Write-Host "Recorded $($output.Count) window(s) to $Path" -ForegroundColor Green
    }
} catch {
    Write-Error $_
    exit 1
}
#endregion
