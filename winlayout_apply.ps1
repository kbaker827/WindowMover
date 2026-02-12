<#
.SYNOPSIS
    Applies saved window layouts to restore window positions and sizes.

.DESCRIPTION
    Reads a JSON configuration file created by winlayout_record.ps1 and moves
    windows to their saved positions. Supports multiple monitors, DPI scaling,
    preset layouts, and automatic process launching.

.PARAMETER Path
    Path to the JSON layout config file. Default: "$env:USERPROFILE\windowlayout.config"

.PARAMETER Process
    When used with -Record, specifies which processes to capture.

.PARAMETER Record
    Record current layout instead of applying. Use with -Process.

.PARAMETER RecordBundleName
    When recording, save as a named bundle in the config file.

.PARAMETER BundleToApply
    Apply a specific named bundle from the config file.

.PARAMETER DryRun
    Show what would be done without actually moving windows.

.EXAMPLE
    .\winlayout_apply.ps1
    Apply the default saved layout.

.EXAMPLE
    .\winlayout_apply.ps1 -Path "dev-setup.json"
    Apply a specific layout file.

.EXAMPLE
    .\winlayout_apply.ps1 -BundleToApply "morning-setup"
    Apply a named bundle from the config.

.EXAMPLE
    .\winlayout_apply.ps1 -Record -Process "chrome","code","notepad" -RecordBundleName "dev"
    Record current positions of Chrome, VS Code, and Notepad as bundle "dev".

.NOTES
    Version: 2.1
    Requires: Windows 8+ with PowerShell 5.1+
    
    Exit Codes:
    0 - Success
    1 - Error
#>

[CmdletBinding()]
param(
    [string]$Path = (Join-Path $env:USERPROFILE 'windowlayout.config'),
    [string[]]$Process,
    [switch]$Record,
    [string]$RecordBundleName,
    [string]$BundleToApply,
    [switch]$DryRun
)

#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$script:DryRun = $DryRun.IsPresent

#region Win32 Interop
$null = @'
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left; public int Top; public int Right; public int Bottom;
}

public class Win32 {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll", EntryPoint="GetDpiForWindow")]
    public static extern uint GetDpiForWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}

public class Shcore {
    [DllImport("shcore.dll")]
    public static extern int GetDpiForSystem();
}
'@ | Add-Type -PassThru -ErrorAction SilentlyContinue
#endregion

#region Constants
$script:SWP = @{
    NOSIZE = 0x0001; NOMOVE = 0x0002; NOZORDER = 0x0004; NOREDRAW = 0x0008
    NOACTIVATE = 0x0010; FRAMECHANGED = 0x0020; SHOWWINDOW = 0x0040
    HIDEWINDOW = 0x0080; NOCOPYBITS = 0x0100; NOOWNERZORDER = 0x0200
}
$script:HWND = @{ TOP = [IntPtr]::Zero; TOPMOST = [IntPtr](-1); NOTOPMOST = [IntPtr](-2); BOTTOM = [IntPtr]1 }

$script:Presets = @{
    Full            = @{ Anchor = 'TopLeft'; WidthPct = 100; HeightPct = 100 }
    LeftHalf        = @{ Anchor = 'Left'; WidthPct = 50; HeightPct = 100 }
    RightHalf       = @{ Anchor = 'Right'; WidthPct = 50; HeightPct = 100 }
    TopHalf         = @{ Anchor = 'Top'; WidthPct = 100; HeightPct = 50 }
    BottomHalf      = @{ Anchor = 'Bottom'; WidthPct = 100; HeightPct = 50 }
    LeftThird       = @{ Anchor = 'TopLeft'; WidthPct = 33.333; HeightPct = 100 }
    CenterThird     = @{ Anchor = 'Top'; WidthPct = 33.333; HeightPct = 100 }
    RightThird      = @{ Anchor = 'TopRight'; WidthPct = 33.333; HeightPct = 100 }
    LeftTwoThirds   = @{ Anchor = 'Left'; WidthPct = 66.666; HeightPct = 100 }
    RightTwoThirds  = @{ Anchor = 'Right'; WidthPct = 66.666; HeightPct = 100 }
    TopLeftQuarter  = @{ Grid = '2x2'; Cell = '1,1' }
    TopRightQuarter = @{ Grid = '2x2'; Cell = '1,2' }
    BottomLeftQuarter = @{ Grid = '2x2'; Cell = '2,1' }
    BottomRightQuarter = @{ Grid = '2x2'; Cell = '2,2' }
    CenteredLarge   = @{ Anchor = 'Center'; WidthPct = 70; HeightPct = 70 }
}
#endregion

#region Helper Functions
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Get-Screens { return [System.Windows.Forms.Screen]::AllScreens }

function Get-WorkingArea {
    param([int]$MonitorIndex = -1)
    $screens = Get-Screens
    if ($MonitorIndex -ge 0 -and $MonitorIndex -lt $screens.Count) {
        return $screens[$MonitorIndex].WorkingArea
    }
    return ($screens | Where-Object { $_.Primary }).WorkingArea
}

function Get-MonitorFromHandle {
    param([IntPtr]$Handle)
    $target = [System.Windows.Forms.Screen]::FromHandle($Handle)
    $screens = Get-Screens
    for ($i = 0; $i -lt $screens.Count; $i++) {
        if ($screens[$i].DeviceName -eq $target.DeviceName) { return $i }
    }
    return -1
}

function Get-WindowHandles {
    param([string]$ProcessName, [string]$TitlePattern)
    
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | 
        Where-Object { $_.MainWindowHandle -ne 0 }
    
    if ($TitlePattern) {
        $procs = $procs | Where-Object { $_.MainWindowTitle -match $TitlePattern }
    }
    
    $results = foreach ($p in $procs) {
        $h = [IntPtr]$p.MainWindowHandle
        if ($h -ne [IntPtr]::Zero -and [Win32]::IsWindowVisible($h)) {
            [PSCustomObject]@{ Process = $p; Handle = $h }
        }
    }
    return $results | Sort-Object { $_.Process.StartTime } -Descending
}

function Expand-EnvVars { param([string]$s) return [Environment]::ExpandEnvironmentVariables($s) }

function Is-ProcessRunning { param([string]$Name) return [bool](Get-Process -Name $Name -ErrorAction SilentlyContinue) }

function Invoke-Launch {
    param([PSCustomObject]$Entry)
    
    if (-not ($Entry.ensureRunning -or $Entry.launchPath)) { return }
    if (Is-ProcessRunning -Name $Entry.processName) { return }
    
    $launchPath = Expand-EnvVars $Entry.launchPath
    if (-not $launchPath) {
        Write-Warning "ensureRunning for '$($Entry.processName)' but no launchPath provided"
        return
    }
    
    Write-Host "[LAUNCH] $($Entry.processName) -> $launchPath" -ForegroundColor Cyan
    if ($script:DryRun) { Write-Host "[DRYRUN] Launch skipped"; return }
    
    $params = @{
        FilePath = $launchPath
        WorkingDirectory = (Expand-EnvVars ($Entry.launchWorkingDir ?? (Split-Path $launchPath)))
        ErrorAction = 'Stop'
    }
    if ($Entry.launchArgs) { $params['ArgumentList'] = $Entry.launchArgs }
    if (($Entry.launchAsUser -eq 'elevated')) { $params['Verb'] = 'runas' }
    
    try { Start-Process @params } catch { throw "Launch failed: $($_.Exception.Message)" }
    
    $delay = [int]($Entry.postLaunchDelaySeconds ?? 0)
    if ($delay -gt 0) { Start-Sleep -Seconds $delay }
}

function Wait-ForWindow {
    param([string]$ProcessName, [string]$TitlePattern, [int]$Retries = 0, [int]$Delay = 1)
    for ($i = 0; $i -le $Retries; $i++) {
        $found = Get-WindowHandles -ProcessName $ProcessName -TitlePattern $TitlePattern | Select-Object -First 1
        if ($found) { return $found }
        if ($i -lt $Retries) { Start-Sleep -Seconds $Delay }
    }
    return $null
}

function Get-Dpi {
    param([IntPtr]$Handle)
    try { return [int][Win32]::GetDpiForWindow($Handle) } catch { return $null }
}

function Get-SystemDpi { try { return [int][Shcore]::GetDpiForSystem() } catch { return 96 } }

function Adjust-RectForDpi {
    param([PSCustomObject]$Rect, [IntPtr]$Handle, [string]$Mode = 'auto')
    
    if ($Mode -eq 'logical') { return $Rect }
    
    $dpi = if ($Mode -eq 'auto') { Get-Dpi -Handle $Handle } else { (Get-Dpi -Handle $Handle) ?? (Get-SystemDpi) }
    if (-not $dpi -or [Math]::Abs($dpi - 96) -lt 0.01) { return $Rect }
    
    $scale = $dpi / 96.0
    return [PSCustomObject]@{
        X = [int][Math]::Round($Rect.X * $scale)
        Y = [int][Math]::Round($Rect.Y * $scale)
        Width = [int][Math]::Round($Rect.Width * $scale)
        Height = [int][Math]::Round($Rect.Height * $scale)
    }
}

function Clamp-Rect {
    param([System.Drawing.Rectangle]$Area, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    return [PSCustomObject]@{
        X = [Math]::Max($Area.Left, [Math]::Min($X, $Area.Right - 50))
        Y = [Math]::Max($Area.Top, [Math]::Min($Y, $Area.Bottom - 50))
        Width = [Math]::Min([Math]::Max(50, $Width), $Area.Width)
        Height = [Math]::Min([Math]::Max(50, $Height), $Area.Height)
    }
}

function Apply-Pad { param([System.Drawing.Rectangle]$Area, [int]$Pad = 0) 
    if ($Pad -le 0) { return $Area }
    return [System.Drawing.Rectangle]::FromLTRB($Area.Left + $Pad, $Area.Top + $Pad, $Area.Right - $Pad, $Area.Bottom - $Pad)
}

function Parse-Grid { param([string]$Text) 
    if ($Text -match '(\d+)\s*x\s*(\d+)') { 
        return @{ Rows = [int]$matches[1]; Cols = [int]$matches[2] }
    }
    return $null
}

function Compute-FromGrid {
    param([System.Drawing.Rectangle]$Area, [PSCustomObject]$Entry)
    
    $g = Parse-Grid -Text $Entry.grid
    if (-not $g) { return $null }
    
    if (-not ($Entry.cell -match '(\d+)\s*,\s*(\d+)')) { throw "Invalid cell format. Use 'row,col'" }
    $row = [int]$matches[1]; $col = [int]$matches[2]
    $rowSpan = [int]($Entry.rowSpan ?? 1); $colSpan = [int]($Entry.colSpan ?? 1)
    $gutter = [int]($Entry.gutter ?? 0); $outer = [int]($Entry.outerGutter ?? 0)
    
    $a = if ($outer -gt 0) { Apply-Pad -Area $Area -Pad $outer } else { $Area }
    $cellW = [double]($a.Width - ($g.Cols - 1) * $gutter) / $g.Cols
    $cellH = [double]($a.Height - ($g.Rows - 1) * $gutter) / $g.Rows
    
    $spanW = [int]([Math]::Round($cellW * $colSpan + $gutter * ($colSpan - 1)))
    $spanH = [int]([Math]::Round($cellH * $rowSpan + $gutter * ($rowSpan - 1)))
    $x = [int]([Math]::Round($a.Left + ($col - 1) * ($cellW + $gutter)))
    $y = [int]([Math]::Round($a.Top + ($row - 1) * ($cellH + $gutter)))
    
    return Clamp-Rect -Area $a -X $x -Y $y -Width $spanW -Height $spanH
}

function Compute-FromAnchor {
    param([System.Drawing.Rectangle]$Area, [PSCustomObject]$Entry)
    
    $w = if ($Entry.widthPct) { [int]($Entry.widthPct / 100.0 * $Area.Width) } 
         elseif ($Entry.width) { [int]$Entry.width } else { $Area.Width }
    $h = if ($Entry.heightPct) { [int]($Entry.heightPct / 100.0 * $Area.Height) }
         elseif ($Entry.height) { [int]$Entry.height } else { $Area.Height }
    
    $anc = ($Entry.anchor ?? 'TopLeft').ToString()
    switch -Regex ($anc) {
        '^(TopLeft|TL)$' { $x = $Area.Left; $y = $Area.Top }
        '^(Top|T)$' { $x = $Area.Left + ($Area.Width - $w) / 2; $y = $Area.Top }
        '^(TopRight|TR)$' { $x = $Area.Right - $w; $y = $Area.Top }
        '^(Left|L)$' { $x = $Area.Left; $y = $Area.Top + ($Area.Height - $h) / 2 }
        '^(Center|C|Middle)$' { $x = $Area.Left + ($Area.Width - $w) / 2; $y = $Area.Top + ($Area.Height - $h) / 2 }
        '^(Right|R)$' { $x = $Area.Right - $w; $y = $Area.Top + ($Area.Height - $h) / 2 }
        '^(BottomLeft|BL)$' { $x = $Area.Left; $y = $Area.Bottom - $h }
        '^(Bottom|B)$' { $x = $Area.Left + ($Area.Width - $w) / 2; $y = $Area.Bottom - $h }
        '^(BottomRight|BR)$' { $x = $Area.Right - $w; $y = $Area.Bottom - $h }
        default { $x = $Area.Left; $y = $Area.Top }
    }
    return Clamp-Rect -Area $Area -X ([int][Math]::Round($x)) -Y ([int][Math]::Round($y)) -Width $w -Height $h
}

function Compute-Rect {
    param([PSCustomObject]$Entry)
    
    $wa = Get-WorkingArea -MonitorIndex ($Entry.monitorIndex ?? -1)
    $wa = Apply-Pad -Area $wa -Pad ([int]($Entry.pad ?? 0))
    
    # Try grid first
    $grid = Compute-FromGrid -Area $wa -Entry $Entry
    if ($grid) { return $grid }
    
    # Then anchor
    if ($Entry.anchor) { return Compute-FromAnchor -Area $wa -Entry $Entry }
    
    # Explicit values
    $w = if ($Entry.widthPct) { [int]($Entry.widthPct / 100.0 * $wa.Width) }
         elseif ($Entry.width) { [int]$Entry.width } else { $wa.Width }
    $h = if ($Entry.heightPct) { [int]($Entry.heightPct / 100.0 * $wa.Height) }
         elseif ($Entry.height) { [int]$Entry.height } else { $wa.Height }
    $x = if ($Entry.xPct) { [int]($wa.Left + $Entry.xPct / 100.0 * $wa.Width) }
         elseif ($Entry.x -ne $null) { [int]($wa.Left + $Entry.x) } else { $wa.Left }
    $y = if ($Entry.yPct) { [int]($wa.Top + $Entry.yPct / 100.0 * $wa.Height) }
         elseif ($Entry.y -ne $null) { [int]($wa.Top + $Entry.y) } else { $wa.Top }
    
    return Clamp-Rect -Area $wa -X $x -Y $y -Width $w -Height $h
}

function Apply-Preset($Entry) {
    if (-not $Entry.preset) { return $Entry }
    $key = $Entry.preset.ToString()
    if (-not $script:Presets.ContainsKey($key)) { throw "Unknown preset: $key" }
    
    $merged = [ordered]@{}
    foreach ($k in $script:Presets[$key].Keys) { $merged[$k] = $script:Presets[$key][$k] }
    foreach ($p in $Entry.PSObject.Properties) { $merged[$p.Name] = $p.Value }
    return [PSCustomObject]$merged
}

function Move-Window {
    param([IntPtr]$Handle, [PSCustomObject]$Rect, [PSCustomObject]$Entry)
    
    # Restore window if minimized
    [void][Win32]::ShowWindow($Handle, 9) # SW_RESTORE
    
    $phys = Adjust-RectForDpi -Rect $Rect -Handle $Handle -Mode ($Entry.dpiMode ?? 'auto')
    
    if ($script:DryRun) {
        Write-Host "[DRYRUN] Would move to ($($phys.X),$($phys.Y),$($phys.Width),$($phys.Height))" -ForegroundColor Yellow
        return
    }
    
    if ($Entry.useSetWindowPos) {
        $flags = 0
        foreach ($f in ($Entry.setWindowPosFlags ?? @('NOZORDER','NOACTIVATE'))) {
            $key = $f.ToString().ToUpper()
            if ($script:SWP.ContainsKey($key)) { $flags = $flags -bor $script:SWP[$key] }
        }
        $insert = $script:HWND[($Entry.zOrder ?? 'TOP')]
        $ok = [Win32]::SetWindowPos($Handle, $insert, $phys.X, $phys.Y, $phys.Width, $phys.Height, [uint32]$flags)
        if (-not $ok) { throw "SetWindowPos failed (Win32: $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))" }
    } else {
        $ok = [Win32]::MoveWindow($Handle, $phys.X, $phys.Y, $phys.Width, $phys.Height, $true)
        if (-not $ok) { throw "MoveWindow failed (Win32: $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))" }
    }
}

function Apply-Entry {
    param([PSCustomObject]$Entry, [PSCustomObject]$Defaults)
    
    # Merge defaults
    $merged = if ($Defaults) {
        $m = [ordered]@{}
        foreach ($p in $Defaults.PSObject.Properties) { $m[$p.Name] = $p.Value }
        foreach ($p in $Entry.PSObject.Properties) { $m[$p.Name] = $p.Value }
        [PSCustomObject]$m
    } else { $Entry }
    
    $merged = Apply-Preset -Entry $merged
    $procName = $merged.processName
    if (-not $procName) { throw "Entry missing processName" }
    
    # Wait before targeting
    $wait = [int]($merged.waitForSeconds ?? 0)
    if ($wait -gt 0) { Write-Host "[WAIT] ${wait}s before targeting '$procName'" -ForegroundColor DarkGray; Start-Sleep -Seconds $wait }
    
    # Launch if needed
    Invoke-Launch -Entry $merged
    
    # Wait for window
    $timeout = [int]($merged.launchTimeoutSeconds ?? 0)
    $delay = [int]($merged.retryDelaySeconds ?? 1)
    $retries = if ($timeout -gt 0 -and $delay -gt 0) { [int][Math]::Ceiling($timeout / $delay) } else { [int]($merged.retryCount ?? 0) }
    
    $target = Wait-ForWindow -ProcessName $procName -TitlePattern $merged.windowTitlePattern -Retries $retries -Delay $delay
    if (-not $target) { Write-Host "[SKIP] '$procName' no visible window after waiting" -ForegroundColor Yellow; return }
    
    # Compute and apply
    $rect = Compute-Rect -Entry $merged
    $mon = $merged.monitorIndex ?? -1
    $area = if ($mon -ge 0) { "monitor#$mon" } else { "primary" }
    
    Write-Host "[MOVE] $procName (PID $($target.Process.Id)) -> ($($rect.X),$($rect.Y),$($rect.Width),$($rect.Height)) on $area" -ForegroundColor Green
    Move-Window -Handle $target.Handle -Rect $rect -Entry $merged
    
    # Verify
    $ver = New-Object RECT
    [void][Win32]::GetWindowRect($target.Handle, [ref]$ver)
    Write-Host "[OK]   Now at ($($ver.Left),$($ver.Top))-($($ver.Right),$($ver.Bottom))" -ForegroundColor DarkGray
}

function Get-LayoutConfig($Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Config not found: $Path" }
    return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Resolve-Bundles($Config, $BundleName) {
    if ($BundleName) {
        if (-not $Config.bundles.$BundleName) { throw "Bundle '$BundleName' not found" }
        return @([PSCustomObject]@{ Defaults = $Config.bundleDefaults; Items = $Config.bundles.$BundleName })
    }
    
    if ($Config -is [System.Collections.IEnumerable] -and -not $Config.bundles) {
        return @([PSCustomObject]@{ Defaults = $null; Items = $Config })
    }
    
    $groups = [System.Collections.Generic.List[object]]::new()
    $defaults = $Config.bundleDefaults
    
    if ($Config.applyBundles) {
        foreach ($name in $Config.applyBundles) {
            if (-not $Config.bundles.$name) { Write-Warning "Missing bundle: $name"; continue }
            $groups.Add([PSCustomObject]@{ Defaults = $defaults; Items = $Config.bundles.$name })
        }
    }
    if ($Config.entries) {
        $groups.Add([PSCustomObject]@{ Defaults = $defaults; Items = $Config.entries })
    }
    
    return $groups
}

function Get-WindowRect($Handle) {
    $r = New-Object RECT
    [void][Win32]::GetWindowRect($Handle, [ref]$r)
    return [PSCustomObject]@{ Left = $r.Left; Top = $r.Top; Right = $r.Right; Bottom = $r.Bottom; Width = $r.Right - $r.Left; Height = $r.Bottom - $r.Top }
}

function Record-Layout {
    param([string[]]$Processes, [string]$Path, [string]$BundleName)
    
    if (-not $Processes) { throw "Specify -Process names to record" }
    
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $Processes) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        if (-not $procs) { Write-Warning "No visible window for: $name"; continue }
        
        foreach ($p in $procs) {
            $h = [IntPtr]$p.MainWindowHandle
            if (-not [Win32]::IsWindowVisible($h)) { continue }
            
            $rect = Get-WindowRect -Handle $h
            $idx = Get-MonitorFromHandle -Handle $h
            $wa = Get-WorkingArea -MonitorIndex $idx
            
            $entries.Add([ordered]@{
                processName = $p.ProcessName
                monitorIndex = $idx
                x = [int]($rect.Left - $wa.Left)
                y = [int]($rect.Top - $wa.Top)
                width = [int]$rect.Width
                height = [int]$rect.Height
                windowTitlePattern = [regex]::Escape($p.MainWindowTitle)
                dpiMode = 'physical'
            })
        }
    }
    
    if ($entries.Count -eq 0) { throw "Nothing recorded. Are the apps open with visible windows?" }
    
    $existing = if (Test-Path -LiteralPath $Path) { 
        try { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json } catch { $null }
    } else { $null }
    
    if ($BundleName) {
        if (-not $existing) { $existing = [PSCustomObject]@{ bundles = @{}; applyBundles = @(); entries = @() } }
        if (-not $existing.bundles) { $existing | Add-Member -NotePropertyName bundles -NotePropertyValue @{} -Force }
        $existing.bundles | Add-Member -NotePropertyName $BundleName -NotePropertyValue $entries -Force
        if ($existing.applyBundles -notcontains $BundleName) { $existing.applyBundles += $BundleName }
        $output = $existing
    } else {
        if ($existing -and $existing.entries) { $existing.entries += $entries; $output = $existing }
        elseif ($existing -and $existing -is [System.Collections.IEnumerable]) { $output = $existing + $entries }
        else { $output = $entries }
    }
    
    $output | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding UTF8
    Write-Host "[RECORDED] $($entries.Count) entries to $Path" -ForegroundColor Green
    if ($BundleName) { Write-Host "[BUNDLE] Created/updated bundle: $BundleName" -ForegroundColor Green }
}
#endregion

#region Main
try {
    if ($Record) {
        Record-Layout -Processes $Process -Path $Path -BundleName $RecordBundleName
        exit 0
    }
    
    $config = Get-LayoutConfig -Path $Path
    $groups = Resolve-Bundles -Config $config -BundleName $BundleToApply
    
    foreach ($group in $groups) {
        foreach ($entry in $group.Items) {
            try { Apply-Entry -Entry $entry -Defaults $group.Defaults }
            catch { Write-Warning "[FAIL] $($entry.processName ?? '?'): $($_.Exception.Message)" }
        }
    }
} catch {
    Write-Error $_
    exit 1
}
#endregion
