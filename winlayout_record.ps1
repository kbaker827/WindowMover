<# 
.SYNOPSIS
  Records current positions of visible top-level windows into JSON (and optional CSV).
  Supports deduplication (largest by area) with per-monitor modes and monitor key selection.

.PARAMETER Path
  Destination JSON file. Default: "$env:USERPROFILE\windowlayout.config"

.PARAMETER CsvPath
  Optional CSV output path written in parallel to JSON.

.PARAMETER ProcessName
  Only include these process names (e.g., "chrome","notepad","Code").

.PARAMETER ExcludeProcessName
  Exclude these process names.

.PARAMETER Append
  Append captured entries to JSON if it exists (tries to merge arrays).

.PARAMETER IncludeMinimized
  Include minimized windows (default: false).

.PARAMETER Deduplicate
  Keep only the largest (by area) window per group.

.PARAMETER DedupBy
  Grouping for deduplication. One of:
    - "process"                   → largest per process
    - "process+title"             → largest per process+title
    - "monitor"                   → largest per monitor (any process)
    - "process+monitor"           → largest per process on each monitor
    - "process+title+monitor"     → largest per process+title on each monitor
  Default: "process" (only used when -Deduplicate is present).

.PARAMETER DedupMonitorBy
  For *per-monitor* dedup modes, choose the monitor key:
    - device (default): groups by monitor_device (e.g., "\\.\DISPLAY2")
    - index          : groups by monitor_index (1-based order)
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

# ----- Win32 interop & helpers -----
try { [void][WinRec_WindowEnum] } catch {
  Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public static class WinRec_Native {
  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern int GetWindowTextLength(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

  [DllImport("user32.dll")]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd); // minimized

  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}

[StructLayout(LayoutKind.Sequential)]
public struct RECT {
  public int Left;
  public int Top;
  public int Right;
  public int Bottom;
}

public static class WinRec_WindowEnum {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

  public static List<IntPtr> GetAllTopWindows() {
    var list = new List<IntPtr>(256);
    EnumWindows((hWnd, lParam) => { list.Add(hWnd); return true; }, IntPtr.Zero);
    return list;
  }

  public static string GetTitle(IntPtr hWnd) {
    int len = WinRec_Native.GetWindowTextLength(hWnd);
    if (len <= 0) return string.Empty;
    var sb = new StringBuilder(len + 1);
    WinRec_Native.GetWindowText(hWnd, sb, sb.Capacity);
    return sb.ToString();
  }

  public static bool GetRect(IntPtr hWnd, out RECT r) {
    return WinRec_Native.GetWindowRect(hWnd, out r);
  }
}
"@
}

Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Get-MonitorInfoFromHandle {
  param([intptr]$Handle)

  $screen = [System.Windows.Forms.Screen]::FromHandle($Handle)
  $index  = [array]::IndexOf([System.Windows.Forms.Screen]::AllScreens, $screen) + 1
  return @{
    device = $screen.DeviceName
    index  = $index
    bounds = $screen.Bounds
    work   = $screen.WorkingArea
  }
}

function Get-WindowEntries {
  param(
    [string[]]$OnlyProcesses,
    [string[]]$ExcludeProcesses,
    [switch]$IncludeMinimized
  )

  $only = if ($OnlyProcesses) { $OnlyProcesses | ForEach-Object { $_.ToLower() } } else { @() }
  $excl = if ($ExcludeProcesses) { $ExcludeProcesses | ForEach-Object { $_.ToLower() } } else { @() }

  $result = New-Object System.Collections.Generic.List[object]
  $all = [WinRec_WindowEnum]::GetAllTopWindows()

  foreach ($h in $all) {
    if (-not [WinRec_Native]::IsWindowVisible($h)) { continue }

    $title = [WinRec_WindowEnum]::GetTitle($h)
    if ([string]::IsNullOrWhiteSpace($title)) { continue }

    # Owning process
    $pid = 0
    [void][WinRec_Native]::GetWindowThreadProcessId($h, [ref]$pid)
    if ($pid -eq 0) { continue }

    try { $p = Get-Process -Id $pid -ErrorAction Stop } catch { continue }
    $pname = $p.ProcessName

    # Filters
    if ($only.Count -gt 0 -and ($only -notcontains $pname.ToLower())) { continue }
    if ($excl.Count -gt 0 -and ($excl -contains $pname.ToLower())) { continue }

    # Bounds
    $rect = New-Object RECT
    if (-not [WinRec_WindowEnum]::GetRect($h, [ref]$rect)) { continue }

    $w = $rect.Right - $rect.Left
    $hgt = $rect.Bottom - $rect.Top
    if ($w -le 0 -or $hgt -le 0) { continue }

    if (-not $IncludeMinimized -and [WinRec_Native]::IsIconic([intptr]$h)) { continue }

    # Monitor info
    $mon = Get-MonitorInfoFromHandle -Handle ([intptr]$h)

    $result.Add([pscustomobject]@{
      processname    = $pname
      title          = $title
      x              = [int]$rect.Left
      y              = [int]$rect.Top
      width          = [int]$w
      height         = [int]$hgt
      area           = [int]($w * $hgt)   # helper for dedup
      monitor_index  = [int]$mon.index
      monitor_device = [string]$mon.device
    }) | Out-Null
  }

  # Stable ordering (by monitor, then process, then title)
  return $result | Sort-Object monitor_index, processname, title
}

function Dedup-Entries {
  param(
    [System.Collections.IEnumerable]$Entries,
    [ValidateSet('process','process+title','monitor','process+monitor','process+title+monitor')]
    [string]$By = 'process',
    [ValidateSet('device','index')]
    [string]$MonitorKey = 'device'
  )
  if (-not $Entries) { return @() }

  $monitorSelector = if ($MonitorKey -eq 'index') { { $_.monitor_index } } else { { $_.monitor_device } }

  switch ($By) {
    'process' {
      $groups = $Entries | Group-Object processname
    }
    'process+title' {
      $groups = $Entries | Group-Object @{Expression = { '{0}||{1}' -f $_.processname, $_.title }}
    }
    'monitor' {
      $groups = $Entries | Group-Object -Property $monitorSelector
    }
    'process+monitor' {
      $groups = $Entries | Group-Object @{Expression = {
        '{0}||{1}' -f $_.processname, (& $monitorSelector.InvokeReturnAsIs($_))
      }}
    }
    'process+title+monitor' {
      $groups = $Entries | Group-Object @{Expression = {
        '{0}||{1}||{2}' -f $_.processname, $_.title, (& $monitorSelector.InvokeReturnAsIs($_))
      }}
    }
  }

  $kept = foreach ($g in $groups) {
    $g.Group | Sort-Object area -Descending | Select-Object -First 1
  }
  return $kept
}

function Write-CsvIfRequested {
  param(
    [System.Collections.IEnumerable]$Entries,
    [string]$CsvPath
  )
  if (-not $CsvPath) { return }
  try {
    $Entries |
      Select-Object processname,title,x,y,width,height,monitor_index,monitor_device |
      Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written: $CsvPath"
  } catch {
    Write-Warning "CSV write failed: $($_.Exception.Message)"
  }
}

# ----- MAIN -----
try {
  $entries = Get-WindowEntries -OnlyProcesses $ProcessName -ExcludeProcesses $ExcludeProcessName -IncludeMinimized:$IncludeMinimized
  if (-not $entries -or $entries.Count -eq 0) {
    Write-Warning "No matching windows found to record."
    return
  }

  if ($Deduplicate) {
    $before = $entries.Count
    $entries = Dedup-Entries -Entries $entries -By $DedupBy -MonitorKey $DedupMonitorBy
    Write-Host ("Deduplicated {0} → {1} entries by {2} (monitor={3})" -f $before, $entries.Count, $DedupBy, $DedupMonitorBy)
  }

  # Always drop the helper 'area' before output
  $outEntries = $entries | Select-Object processname,title,x,y,width,height,monitor_index,monitor_device

  # CSV (optional)
  Write-CsvIfRequested -Entries $outEntries -CsvPath $CsvPath

  $json = $outEntries | ConvertTo-Json -Depth 5

  if ($Append -and (Test-Path -LiteralPath $Path)) {
    try {
      $existing = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
      $merged = if ($existing -is [System.Collections.IEnumerable]) { @($existing) + @($outEntries) } else { @($existing) + @($outEntries) }
      $merged | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
      Write-Host "Appended $(($outEntries).Count) entries to $Path"
    } catch {
      Write-Warning "Append failed (invalid existing JSON?). Writing a new file instead. Error: $($_.Exception.Message)"
      $json | Set-Content -LiteralPath $Path -Encoding UTF8
      Write-Host "Wrote fresh file to $Path"
    }
  } else {
    if (Test-Path -LiteralPath $Path) {
      $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
      $bak = "$Path.bak-$stamp"
      Copy-Item -LiteralPath $Path -Destination $bak -ErrorAction SilentlyContinue
      Write-Host "Backup created: $bak"
    }
    $json | Set-Content -LiteralPath $Path -Encoding UTF8
    Write-Host "Recorded $($outEntries.Count) window(s) to $Path"
  }
}
catch {
  Write-Error $_
  exit 1
}