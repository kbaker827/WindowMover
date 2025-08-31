param(
  [switch]$Record,
  [string[]]$Process,
  [string]$RecordBundleName,
  [string]$BundleToApply,
  [switch]$DryRun,
  [string]$OutputPath
)

#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:DryRun = $DryRun.IsPresent

# =========================
#  Win32 P/Invoke helpers
# =========================
try { [void][WindowNative] } catch {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct RECT
{
  public int Left;
  public int Top;
  public int Right;
  public int Bottom;
}

public static class WindowNative
{
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
}

public static class ShcoreNative
{
  [DllImport("shcore.dll")]
  public static extern int GetDpiForSystem();
}
"@
}

# =========================
#   Config & Utilities
# =========================
$ConfigPath = Join-Path $env:USERPROFILE 'windowlayout.config'
if (-not $OutputPath) { $OutputPath = $ConfigPath }

function Read-WindowLayoutConfig {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Config file not found: $Path" }
    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if (-not $json) { throw "Config file is empty or invalid JSON: $Path" }
    $json
}

function Get-Screens { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Screen]::AllScreens }
function Get-WorkingArea { param([int]$MonitorIndex=-1)
    $screens = Get-Screens
    if ($MonitorIndex -ge 0 -and $MonitorIndex -lt $screens.Count) { return $screens[$MonitorIndex].WorkingArea }
    return ($screens | Where-Object { $_.Primary }).WorkingArea
}
function Get-ScreenIndexFromHandle {
    param([IntPtr]$Handle)
    Add-Type -AssemblyName System.Windows.Forms
    $target = [System.Windows.Forms.Screen]::FromHandle($Handle)
    $all = [System.Windows.Forms.Screen]::AllScreens
    for ($i=0; $i -lt $all.Count; $i++){ if ($all[$i].DeviceName -eq $target.DeviceName) { return $i } }
    return -1
}

function Get-EligibleMainWindowHandles {
    param([Parameter(Mandatory)][string]$ProcessName,[string]$WindowTitlePattern)
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
    if ($WindowTitlePattern) { $procs = $procs | Where-Object { $_.MainWindowTitle -match $WindowTitlePattern } }
    $eligible = foreach ($p in $procs) {
        $h = [IntPtr]$p.MainWindowHandle
        if ($h -ne [IntPtr]::Zero -and [WindowNative]::IsWindowVisible($h)) { [PSCustomObject]@{ Process=$p; Handle=$h } }
    }
    $eligible | Sort-Object { $_.Process.StartTime } -Descending
}

function Ensure-Restored { param([IntPtr]$Handle) [void][WindowNative]::ShowWindow($Handle, 9) } # SW_RESTORE

function Clamp-ToArea {
    param([System.Drawing.Rectangle]$Area,[int]$X,[int]$Y,[int]$Width,[int]$Height)
    $nx = [Math]::Max($Area.Left, [Math]::Min($X, $Area.Right  - 50))
    $ny = [Math]::Max($Area.Top,  [Math]::Min($Y, $Area.Bottom - 50))
    $nw = [Math]::Min([Math]::Max(50, $Width),  $Area.Width)
    $nh = [Math]::Min([Math]::Max(50, $Height), $Area.Height)
    [PSCustomObject]@{ X=[int]$nx; Y=[int]$ny; Width=[int]$nw; Height=[int]$nh }
}
function Get-AreaWithPad { param([System.Drawing.Rectangle]$Area,[int]$Pad=0)
    if ($Pad -le 0) { return $Area }
    [System.Drawing.Rectangle]::FromLTRB($Area.Left+$Pad,$Area.Top+$Pad,$Area.Right-$Pad,$Area.Bottom-$Pad)
}

# -------- Presets (single-window sugar) ----------
$__PresetMap = @{
  'Full'            = @{ anchor='TopLeft'; widthPct=100; heightPct=100 }
  'LeftHalf'        = @{ anchor='Left'; widthPct=50; heightPct=100 }
  'RightHalf'       = @{ anchor='Right'; widthPct=50; heightPct=100 }
  'TopHalf'         = @{ anchor='Top'; widthPct=100; heightPct=50 }
  'BottomHalf'      = @{ anchor='Bottom'; widthPct=100; heightPct=50 }
  'LeftThird'       = @{ anchor='TopLeft'; widthPct=33.333; heightPct=100 }
  'CenterThird'     = @{ anchor='Top'; widthPct=33.333; heightPct=100 }
  'RightThird'      = @{ anchor='TopRight'; widthPct=33.333; heightPct=100 }
  'LeftTwoThirds'   = @{ anchor='Left'; widthPct=66.666; heightPct=100 }
  'RightTwoThirds'  = @{ anchor='Right'; widthPct=66.666; heightPct=100 }
  'TopLeftQuarter'  = @{ grid='2x2'; cell='1,1' }
  'TopRightQuarter' = @{ grid='2x2'; cell='1,2' }
  'BottomLeftQuarter' = @{ grid='2x2'; cell='2,1' }
  'BottomRightQuarter'= @{ grid='2x2'; cell='2,2' }
  'CenteredLarge'   = @{ anchor='Center'; widthPct=70; heightPct=70 }
}
function Apply-PresetToEntry {
    param([pscustomobject]$Entry)
    if (-not $Entry.preset) { return $Entry }
    $key = $Entry.preset.ToString()
    $base = $__PresetMap[$key]; if (-not $base) { throw "Unknown preset '$key'." }
    $merged = [ordered]@{}; foreach ($k in $base.Keys) { $merged[$k]=$base[$k] }; foreach ($p in $Entry.PSObject.Properties){ $merged[$p.Name]=$p.Value }
    [pscustomobject]$merged
}

# -------- Grid & Anchors ----------
function Parse-Grid { param([string]$GridText)
    if (-not $GridText -or $GridText -notmatch '^\s*(\d+)\s*x\s*(\d+)\s*$') { return $null }
    [PSCustomObject]@{ Rows=[int]$Matches[1]; Cols=[int]$Matches[2] }
}
function Compute-FromGrid {
    param([System.Drawing.Rectangle]$Area,[pscustomobject]$Entry)
    $g = Parse-Grid -GridText $Entry.grid
    if (-not $g) { return $null }
    $rows=[Math]::Max(1,$g.Rows); $cols=[Math]::Max(1,$g.Cols)
    if (-not $Entry.cell -or $Entry.cell -notmatch '^\s*(\d+)\s*,\s*(\d+)\s*$') { throw "grid specified but cell missing/invalid (use 'row,col')." }
    $r=[int]$Matches[1]; $c=[int]$Matches[2]
    $rowSpan=[int]($Entry.rowSpan ?? 1); $colSpan=[int]($Entry.colSpan ?? 1)
    if ($r -lt 1 -or $r -gt $rows -or $c -lt 1 -or $c -gt $cols) { throw "cell ($r,$c) out of bounds ($rows x $cols)." }
    $gutter=[int]($Entry.gutter ?? 0); $outer=[int]($Entry.outerGutter ?? 0)
    $A = if ($outer -gt 0){[System.Drawing.Rectangle]::FromLTRB($Area.Left+$outer,$Area.Top+$outer,$Area.Right-$outer,$Area.Bottom-$outer)} else {$Area}
    $cellW=[double]($A.Width-($cols-1)*$gutter)/$cols; $cellH=[double]($A.Height-($rows-1)*$gutter)/$rows
    $spanW=[int]([Math]::Round($cellW*$colSpan + $gutter*($colSpan-1))); $spanH=[int]([Math]::Round($cellH*$rowSpan + $gutter*($rowSpan-1)))
    $x=[int]([Math]::Round($A.Left + ($c-1)*($cellW+$gutter))); $y=[int]([Math]::Round($A.Top + ($r-1)*($cellH+$gutter)))
    [PSCustomObject]@{ X=$x; Y=$y; Width=$spanW; Height=$spanH }
}
function Compute-FromAnchor {
    param([System.Drawing.Rectangle]$Area,[pscustomobject]$Entry)
    $w = if ($Entry.widthPct -ne $null){[int]([double]$Entry.widthPct/100.0*$Area.Width)} elseif ($Entry.width -ne $null){[int]$Entry.width}else{$Area.Width}
    $h = if ($Entry.heightPct -ne $null){[int]([double]$Entry.heightPct/100.0*$Area.Height)} elseif ($Entry.height -ne $null){[int]$Entry.height}else{$Area.Height}
    $anc = ($Entry.anchor ?? 'TopLeft').ToString()
    switch -Regex ($anc) {
        '^(TopLeft|TL)$'        {$x=$Area.Left; $y=$Area.Top}
        '^(Top|T)$'             {$x=$Area.Left+($Area.Width-$w)/2; $y=$Area.Top}
        '^(TopRight|TR)$'       {$x=$Area.Right-$w; $y=$Area.Top}
        '^(Left|L)$'            {$x=$Area.Left; $y=$Area.Top+($Area.Height-$h)/2}
        '^(Center|C|Middle)$'   {$x=$Area.Left+($Area.Width-$w)/2; $y=$Area.Top+($Area.Height-$h)/2}
        '^(Right|R)$'           {$x=$Area.Right-$w; $y=$Area.Top+($Area.Height-$h)/2}
        '^(BottomLeft|BL)$'     {$x=$Area.Left; $y=$Area.Bottom-$h}
        '^(Bottom|B)$'          {$x=$Area.Left+($Area.Width-$w)/2; $y=$Area.Bottom-$h}
        '^(BottomRight|BR)$'    {$x=$Area.Right-$w; $y=$Area.Bottom-$h}
        default                 {$x=$Area.Left; $y=$Area.Top}
    }
    Clamp-ToArea -Area $Area -X ([int][Math]::Round($x)) -Y ([int][Math]::Round($y)) -Width $w -Height $h
}

# -------- DPI helpers ----------
function Get-DpiForHandle { param([IntPtr]$Handle) try { [int][WindowNative]::GetDpiForWindow($Handle) } catch { $null } }
function Get-SystemDpi   { try { [int][ShcoreNative]::GetDpiForSystem() } catch { 96 } }
function Adjust-ForDpi {
    param([pscustomobject]$Rect,[IntPtr]$Handle,[string]$Mode='auto')
    $scale = $null
    if ($Mode -eq 'logical') { return $Rect }
    if ($Mode -eq 'auto')    { $dpi = Get-DpiForHandle -Handle $Handle; if ($dpi) { $scale = $dpi/96.0 } else { return $Rect } }
    if ($Mode -eq 'physical'){ $dpi = (Get-DpiForHandle -Handle $Handle) ?? (Get-SystemDpi); $scale = $dpi/96.0 }
    if (-not $scale -or [Math]::Abs($scale-1.0) -lt 0.01) { return $Rect }
    [pscustomobject]@{
        X=[int][Math]::Round($Rect.X*$scale); Y=[int][Math]::Round($Rect.Y*$scale);
        Width=[int][Math]::Round($Rect.Width*$scale); Height=[int][Math]::Round($Rect.Height*$scale)
    }
}

# -------- Rect planner ----------
function Compute-RectFromEntry {
    param([pscustomobject]$Entry)
    $wa0 = Get-WorkingArea -MonitorIndex ($Entry.monitorIndex ?? -1)
    $wa  = Get-AreaWithPad -Area $wa0 -Pad ([int]($Entry.pad ?? 0))

    # 1) grid
    $gr = Compute-FromGrid -Area $wa -Entry $Entry
    if ($gr) { return (Clamp-ToArea -Area $wa -X $gr.X -Y $gr.Y -Width $gr.Width -Height $gr.Height) }

    # 2) anchor
    if ($Entry.PSObject.Properties.Name -contains 'anchor') { return (Compute-FromAnchor -Area $wa -Entry $Entry) }

    # 3) explicit/percent
    $w = if ($Entry.widthPct -ne $null){[int]([double]$Entry.widthPct/100.0*$wa.Width)} elseif ($Entry.width -ne $null){[int]$Entry.width}else{$wa.Width}
    $h = if ($Entry.heightPct -ne $null){[int]([double]$Entry.heightPct/100.0*$wa.Height)} elseif ($Entry.height -ne $null){[int]$Entry.height}else{$wa.Height}
    $x = if ($Entry.xPct -ne $null){[int]($wa.Left + ([double]$Entry.xPct/100.0*$wa.Width))} elseif ($Entry.x -ne $null){[int]($wa.Left + [int]$Entry.x)} else {$wa.Left}
    $y = if ($Entry.yPct -ne $null){[int]($wa.Top  + ([double]$Entry.yPct/100.0*$wa.Height))} elseif ($Entry.y -ne $null){[int]($wa.Top + [int]$Entry.y)} else {$wa.Top}
    Clamp-ToArea -Area $wa -X $x -Y $y -Width $w -Height $h
}

# -------- Move via MoveWindow or SetWindowPos ----------
$__SWP = @{
  'NOSIZE'=0x0001; 'NOMOVE'=0x0002; 'NOZORDER'=0x0004; 'NOREDRAW'=0x0008; 'NOACTIVATE'=0x0010;
  'FRAMECHANGED'=0x0020; 'SHOWWINDOW'=0x0040; 'HIDEWINDOW'=0x0080; 'NOCOPYBITS'=0x0100;
  'NOOWNERZORDER'=0x0200; 'NOSENDCHANGING'=0x0400; 'ASYNCWINDOWPOS'=0x4000
}
$__HWND = @{ 'TOP'=[IntPtr]::Zero; 'TOPMOST'=[IntPtr](-1); 'NOTOPMOST'=[IntPtr](-2); 'BOTTOM'=[IntPtr]1 }

function Move-WindowSmart {
    param([IntPtr]$Handle,[pscustomobject]$Rect,[pscustomobject]$Entry)
    Ensure-Restored -Handle $Handle

    $dpiMode = ($Entry.dpiMode ?? 'auto').ToString().ToLower()
    $physRect = Adjust-ForDpi -Rect $Rect -Handle $Handle -Mode $dpiMode

    if ($script:DryRun) {
        Write-Host "[DRYRUN] would move -> ($($physRect.X),$($physRect.Y),$($physRect.Width),$($physRect.Height))"
        return
    }

    if ($Entry.useSetWindowPos) {
        $flags = 0
        foreach ($f in ($Entry.setWindowPosFlags ?? @('NOZORDER','NOACTIVATE'))) {
            $name = $f.ToString().ToUpper()
            if ($__SWP.ContainsKey($name)) { $flags = $flags -bor $__SWP[$name] }
        }
        $insertAfter = $__HWND[ ($Entry.zOrder ?? 'TOP') ]
        $ok = [WindowNative]::SetWindowPos($Handle,$insertAfter,$physRect.X,$physRect.Y,$physRect.Width,$physRect.Height,[uint32]$flags)
        if (-not $ok) { $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error(); throw "SetWindowPos failed [Win32=$code]" }
    } else {
        $ok = [WindowNative]::MoveWindow($Handle,$physRect.X,$physRect.Y,$physRect.Width,$physRect.Height,$true)
        if (-not $ok) { $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error(); throw "MoveWindow failed [Win32=$code]" }
    }
}

# -------- Launch helpers ----------
function Expand-EnvText { param([string]$s) if (-not $s) { return $s }; return [Environment]::ExpandEnvironmentVariables($s) }
function Is-ProcessRunning { param([string]$Name) (Get-Process -Name $Name -ErrorAction SilentlyContinue) -ne $null }

function Launch-IfNeeded {
    param([pscustomobject]$Entry)

    $ensure = [bool]($Entry.ensureRunning ?? $false)
    if (-not $ensure) { return }

    $procName = $Entry.processName
    if (Is-ProcessRunning -Name $procName) { return }

    $path = Expand-EnvText ($Entry.launchPath ?? '')
    if (-not $path) {
        Write-Warning "ensureRunning=true for '$procName' but no launchPath provided; will wait for process if it starts elsewhere."
        return
    }
    $args = $Entry.launchArgs
    $cwd  = Expand-EnvText ($Entry.launchWorkingDir ?? [IO.Path]::GetDirectoryName($path))
    $verb = (($Entry.launchAsUser ?? 'current').ToString().ToLower() -eq 'elevated') ? 'runas' : $null

    Write-Host "[LAUNCH] $procName -> $path $args"
    if ($script:DryRun) { Write-Host "[DRYRUN] (launch skipped)"; return }

    $si = @{
        FilePath     = $path
        WorkingDirectory = $cwd
        ErrorAction  = 'Stop'
    }
    if ($args) { $si['ArgumentList'] = $args }
    if ($verb) { $si['Verb'] = $verb }

    try { Start-Process @si } catch { throw "Failed to launch '$procName' at '$path': $($_.Exception.Message)" }

    $postDelay = [int]($Entry.postLaunchDelaySeconds ?? 0)
    if ($postDelay -gt 0) { Start-Sleep -Seconds $postDelay }
}

# -------- Orchestration ----------
function Wait-ForProcessWindow {
    param([string]$ProcessName,[string]$WindowTitlePattern,[int]$RetryCount=0,[int]$RetryDelaySeconds=1)
    $try = 0
    while ($true) {
        $eligible = Get-EligibleMainWindowHandles -ProcessName $ProcessName -WindowTitlePattern $WindowTitlePattern
        if ($eligible) { return ($eligible | Select-Object -First 1) }
        if ($try -ge $RetryCount) { return $null }
        Start-Sleep -Seconds $RetryDelaySeconds
        $try++
    }
}

function Apply-WindowLayoutEntry {
    param([pscustomobject]$EntryRaw,[pscustomobject]$BundleDefaults)

    # Merge bundle defaults -> entry -> preset
    $EntryMerged = if ($BundleDefaults) {
        $m=[ordered]@{}; foreach($p in $BundleDefaults.PSObject.Properties){$m[$p.Name]=$p.Value}
        foreach($p in $EntryRaw.PSObject.Properties){$m[$p.Name]=$p.Value}
        [pscustomobject]$m
    } else { $EntryRaw }

    $Entry = Apply-PresetToEntry -Entry $EntryMerged
    $procName = $Entry.processName
    if (-not $procName) { throw "Config entry missing 'processName': $($Entry | ConvertTo-Json -Compress)" }

    # Optional initial wait
    $wait = [int]($Entry.waitForSeconds ?? 0)
    if ($wait -gt 0) { Write-Host ("[WAIT] {0}s before targeting '{1}'" -f $wait,$procName); Start-Sleep -Seconds $wait }

    # Launch if needed
    Launch-IfNeeded -Entry $Entry

    # Window appearance waits
    $retryDelay=[int]($Entry.retryDelaySeconds ?? 1)
    $retryCount=[int]($Entry.retryCount ?? 0)

    # If a specific launch timeout is set, override retries to match it
    $launchTimeout = [int]($Entry.launchTimeoutSeconds ?? 0)
    if ($launchTimeout -gt 0 -and $retryDelay -gt 0) {
        $retryCount = [int][Math]::Ceiling($launchTimeout / $retryDelay)
    }

    $target = Wait-ForProcessWindow -ProcessName $procName -WindowTitlePattern $Entry.windowTitlePattern -RetryCount $retryCount -RetryDelaySeconds $retryDelay
    if (-not $target) { Write-Host "[SKIP] '$procName' has no visible main window after retries."; return }

    $rect = Compute-RectFromEntry -Entry $Entry
    $h = $target.Handle; $pid=$target.Process.Id; $title=$target.Process.MainWindowTitle
    $monitorIdx = ($Entry.monitorIndex ?? -1); $area = if ($monitorIdx -ge 0) { "monitor#$monitorIdx" } else { "primary monitor" }
    $rectStr = "({0},{1},{2},{3})" -f $rect.X,$rect.Y,$rect.Width,$rect.Height

    Write-Host "[MOVE] $procName (PID $pid) title='$title' -> $rectStr on $area (dpiMode=${($Entry.dpiMode ?? 'auto')})"
    Move-WindowSmart -Handle $h -Rect $rect -Entry $Entry

    $ver = New-Object RECT
    [void][WindowNative]::GetWindowRect($h,[ref]$ver)
    Write-Host ("[OK]   New rect: ({0},{1})-({2},{3})" -f $ver.Left,$ver.Top,$ver.Right,$ver.Bottom)
}

# -------- Bundles (multi-window presets) ----------
function Resolve-EntriesFromConfig {
    param($Config,[string]$BundleToApply)
    if ($BundleToApply) {
        if (-not $Config.PSObject.Properties.Name -contains 'bundles') { throw "No 'bundles' found in config." }
        $bundle = $Config.bundles.$BundleToApply
        if (-not $bundle) { throw "Bundle '$BundleToApply' not found." }
        $defaults = $Config.bundleDefaults
        return @([pscustomobject]@{ Defaults=$defaults; Items=$bundle })
    }

    if ($Config -is [System.Collections.IEnumerable] -and -not ($Config.PSObject.Properties.Name -contains 'entries')) {
        return @([pscustomobject]@{ Defaults=$null; Items=$Config })
    }

    $all = @()
    $defaults = $Config.bundleDefaults

    if ($Config.PSObject.Properties.Name -contains 'applyBundles') {
        foreach ($name in $Config.applyBundles) {
            $b = $Config.bundles.$name
            if (-not $b) { Write-Warning "applyBundles references missing bundle '$name'"; continue }
            $all += [pscustomobject]@{ Defaults=$defaults; Items=$b }
        }
    }
    if ($Config.PSObject.Properties.Name -contains 'entries' -and $Config.entries) {
        $all += [pscustomobject]@{ Defaults=$defaults; Items=$Config.entries }
    }
    return ,$all
}

# -------- Recorder ----------
function Get-WindowRectObject {
    param([IntPtr]$Handle)
    $r = New-Object RECT
    [void][WindowNative]::GetWindowRect($Handle,[ref]$r)
    [pscustomobject]@{ Left=$r.Left; Top=$r.Top; Right=$r.Right; Bottom=$r.Bottom; Width=($r.Right-$r.Left); Height=($r.Bottom-$r.Top) }
}
function Record-WindowLayout {
    param([string[]]$Processes,[string]$Path,[string]$BundleName)
    if (-not $Processes -or $Processes.Count -eq 0) { throw "Specify one or more -Process names to record (e.g., -Process code,chrome,notepad)." }
    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($name in $Processes) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        if (-not $procs) { Write-Warning "No visible main window for process '$name'."; continue }
        foreach ($p in $procs) {
            $h = [IntPtr]$p.MainWindowHandle
            if (-not [WindowNative]::IsWindowVisible($h)) { continue }
            $rect = Get-WindowRectObject -Handle $h
            $idx  = Get-ScreenIndexFromHandle -Handle $h
            $wa   = Get-WorkingArea -MonitorIndex $idx
            $xRel = [int]($rect.Left - $wa.Left)
            $yRel = [int]($rect.Top  - $wa.Top)
            $entry = [ordered]@{
                processName        = $p.ProcessName
                monitorIndex       = $idx
                x                  = $xRel
                y                  = $yRel
                width              = [int]$rect.Width
                height             = [int]$rect.Height
                windowTitlePattern = [regex]::Escape($p.MainWindowTitle)
                dpiMode            = 'physical'
            }
            [void]$entries.Add([pscustomobject]$entry)
        }
    }
    if ($entries.Count -eq 0) { throw "Nothing recorded. Are those apps open with visible windows?" }

    $existing = $null
    if (Test-Path -LiteralPath $Path) { try { $existing = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json } catch {} }

    if ($BundleName) {
        if (-not $existing) { $existing = [pscustomobject]@{ bundles=@{}; applyBundles=@(); entries=@(); bundleDefaults=@{} } }
        if (-not ($existing.PSObject.Properties.Name -contains 'bundles')) { $existing | Add-Member -NotePropertyName bundles -NotePropertyValue @{} -Force }
        $existing.bundles | Add-Member -NotePropertyName $BundleName -NotePropertyValue $entries -Force
        if (-not ($existing.applyBundles -contains $BundleName)) { $existing.applyBundles += $BundleName }
        $outputObj = $existing
    } else {
        if ($existing -and ($existing.PSObject.Properties.Name -contains 'entries')) { $existing.entries += $entries; $outputObj=$existing }
        elseif ($existing -and $existing -is [System.Collections.IEnumerable]) { $outputObj = $existing + $entries }
        else { $outputObj = $entries }
    }

    $json = $outputObj | ConvertTo-Json -Depth 10
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
    Write-Host "[RECORDED] Wrote $($entries.Count) entries to $Path" -ForegroundColor Green
    if ($BundleName) { Write-Host "[BUNDLE] Bundle '$BundleName' updated/created in config." -ForegroundColor Green }
}

# =========================
#   MAIN
# =========================
try {
    if ($Record) { Record-WindowLayout -Processes $Process -Path $OutputPath -BundleName $RecordBundleName; exit 0 }

    $config = Read-WindowLayoutConfig -Path $OutputPath
    $groups = Resolve-EntriesFromConfig -Config $config -BundleToApply $BundleToApply

    foreach ($g in $groups) {
        foreach ($entry in $g.Items) {
            try { Apply-WindowLayoutEntry -Entry $entry -BundleDefaults $g.Defaults }
            catch { Write-Warning ("[FAIL] {0}: {1}" -f ($entry.processName ?? '<unknown>'), $_.Exception.Message) }
        }
    }
}
catch { Write-Error $_; exit 1 }