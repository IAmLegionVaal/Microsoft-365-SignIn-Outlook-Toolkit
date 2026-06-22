#requires -Version 5.1
<#!
.SYNOPSIS
    Guarded Microsoft Outlook repair toolkit.
.DESCRIPTION
    Repairs common classic Outlook problems while preserving recoverable backups.
    Supports dry-run, confirmation prompts, before/after snapshots and detailed logs.
.NOTES
    Created by Dewald Pretorius - L2 IT Support Engineer.
    PST files are never modified. OST and cache repairs rename data into timestamped
    backup folders instead of deleting it.
#>

[CmdletBinding()]
param(
    [switch]$RepairAllSafe,
    [switch]$RestartOutlook,
    [switch]$RebuildRoamCache,
    [switch]$ResetNavigationPane,
    [switch]$ResetViews,
    [switch]$RebuildOstCache,
    [switch]$DisableUserAddIns,
    [switch]$RestartSearchService,
    [switch]$RestartClickToRun,
    [switch]$FlushDns,
    [switch]$OpenProfileManager,
    [switch]$OpenOfficeRepair,
    [switch]$DryRun,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '2.0.0'
$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Outlook_Repair_Logs'
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $OutputPath "Outlook_Repair_$RunStamp.log"
$BackupRoot = Join-Path $OutputPath "Backup_$RunStamp"
New-Item -Path $BackupRoot -ItemType Directory -Force | Out-Null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DRYRUN')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'    { Write-Host $Message -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        'DRYRUN'  { Write-Host "DRY RUN: $Message" -ForegroundColor Cyan }
        default   { Write-Host $Message }
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Confirm-Repair {
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$HighImpact
    )
    if ($DryRun) { return $true }
    if ($HighImpact) {
        return (Read-Host "$Message Type REPAIR to continue") -eq 'REPAIR'
    }
    return (Read-Host "$Message Type YES to continue") -eq 'YES'
}

function Get-OutlookPath {
    $command = Get-Command OUTLOOK.EXE -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $roots = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "$env:ProgramFiles\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
    )
    foreach ($root in $roots) {
        if (-not [string]::IsNullOrWhiteSpace($root)) {
            $candidate = Join-Path $root 'OUTLOOK.EXE'
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
    }
    return $null
}

function Stop-OutlookProcess {
    $processes = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)
    foreach ($process in $processes) {
        try {
            if ($DryRun) {
                Write-Log "Would close Outlook process ID $($process.Id)." 'DRYRUN'
                continue
            }
            [void]$process.CloseMainWindow()
        } catch {}
    }

    if (-not $DryRun -and $processes.Count -gt 0) {
        Start-Sleep -Seconds 3
        Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Get-OutlookSnapshot {
    param([Parameter(Mandatory)][string]$Stage)

    $profilePath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
    $addInPath = 'HKCU:\Software\Microsoft\Office\Outlook\Addins'
    $outlookData = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
    $roamCache = Join-Path $outlookData 'RoamCache'

    $snapshot = [ordered]@{
        Stage = $Stage
        Generated = (Get-Date).ToString('o')
        Computer = $env:COMPUTERNAME
        User = "$env:USERDOMAIN\$env:USERNAME"
        IsAdministrator = Test-IsAdministrator
        OutlookPath = Get-OutlookPath
        OutlookProcesses = @(
            Get-Process OUTLOOK -ErrorAction SilentlyContinue |
                Select-Object Id, ProcessName, Path, StartTime
        )
        Profiles = @(
            Get-ChildItem -LiteralPath $profilePath -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty PSChildName
        )
        OstFiles = @(
            Get-ChildItem -LiteralPath $outlookData -Filter '*.ost' -File -ErrorAction SilentlyContinue |
                Select-Object FullName, Length, LastWriteTime
        )
        PstFiles = @(
            Get-ChildItem -LiteralPath $outlookData -Filter '*.pst' -File -ErrorAction SilentlyContinue |
                Select-Object FullName, Length, LastWriteTime
        )
        RoamCacheExists = Test-Path -LiteralPath $roamCache
        UserAddIns = @(
            Get-ChildItem -LiteralPath $addInPath -ErrorAction SilentlyContinue | ForEach-Object {
                $item = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    Name = $_.PSChildName
                    LoadBehavior = $item.LoadBehavior
                    FriendlyName = $item.FriendlyName
                }
            }
        )
        Services = @(
            Get-Service ClickToRunSvc, WSearch -ErrorAction SilentlyContinue |
                Select-Object Name, Status, StartType
        )
    }

    $path = Join-Path $OutputPath "Outlook_${Stage}_$RunStamp.json"
    $snapshot | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $path -Encoding UTF8
    Write-Log "Saved $Stage snapshot: $path" 'SUCCESS'
}

function Invoke-RebuildRoamCache {
    $path = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\RoamCache'
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Log 'Outlook RoamCache was not found. No change was needed.' 'WARN'
        return
    }
    if (-not (Confirm-Repair 'Close Outlook and rebuild its RoamCache?')) { return }

    Stop-OutlookProcess
    $destination = Join-Path $BackupRoot 'RoamCache'
    if ($DryRun) {
        Write-Log "Would move $path to $destination." 'DRYRUN'
        return
    }

    Move-Item -LiteralPath $path -Destination $destination -Force
    Write-Log "RoamCache moved to $destination. Outlook will rebuild it." 'SUCCESS'
}

function Invoke-ResetNavigationPane {
    $outlookPath = Get-OutlookPath
    if (-not $outlookPath) {
        Write-Log 'OUTLOOK.EXE was not found.' 'ERROR'
        return
    }
    if (-not (Confirm-Repair 'Reset the Outlook navigation pane?')) { return }

    Stop-OutlookProcess
    if ($DryRun) {
        Write-Log "Would start $outlookPath /resetnavpane." 'DRYRUN'
        return
    }
    Start-Process -FilePath $outlookPath -ArgumentList '/resetnavpane'
    Write-Log 'Outlook navigation pane reset was started.' 'SUCCESS'
}

function Invoke-ResetViews {
    $outlookPath = Get-OutlookPath
    if (-not $outlookPath) {
        Write-Log 'OUTLOOK.EXE was not found.' 'ERROR'
        return
    }
    if (-not (Confirm-Repair 'Reset all custom Outlook folder views? This removes custom views.' -HighImpact)) { return }

    Stop-OutlookProcess
    if ($DryRun) {
        Write-Log "Would start $outlookPath /cleanviews." 'DRYRUN'
        return
    }
    Start-Process -FilePath $outlookPath -ArgumentList '/cleanviews'
    Write-Log 'Outlook view reset was started.' 'SUCCESS'
}

function Invoke-RebuildOstCache {
    $outlookData = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
    $ostFiles = @(Get-ChildItem -LiteralPath $outlookData -Filter '*.ost' -File -ErrorAction SilentlyContinue)
    if ($ostFiles.Count -eq 0) {
        Write-Log 'No OST cache files were found.' 'WARN'
        return
    }

    Write-Host 'OST files selected for backup and rebuild:' -ForegroundColor Yellow
    $ostFiles | Select-Object FullName, @{n='SizeGB';e={[math]::Round($_.Length / 1GB, 2)}}, LastWriteTime | Format-Table -AutoSize
    if (-not (Confirm-Repair 'Rebuild the Outlook OST cache files? Ensure cloud mail is fully synchronised.' -HighImpact)) { return }

    Stop-OutlookProcess
    $ostBackup = Join-Path $BackupRoot 'OST'
    if (-not $DryRun) { New-Item -Path $ostBackup -ItemType Directory -Force | Out-Null }

    foreach ($file in $ostFiles) {
        $destination = Join-Path $ostBackup $file.Name
        if ($DryRun) {
            Write-Log "Would move $($file.FullName) to $destination." 'DRYRUN'
        } else {
            Move-Item -LiteralPath $file.FullName -Destination $destination -Force
            Write-Log "Backed up OST: $($file.Name)" 'SUCCESS'
        }
    }
    if (-not $DryRun) {
        Write-Log 'OST files were backed up. Outlook will recreate them at next start.' 'SUCCESS'
    }
}

function Invoke-DisableUserAddIns {
    $path = 'HKCU:\Software\Microsoft\Office\Outlook\Addins'
    $keys = @(Get-ChildItem -LiteralPath $path -ErrorAction SilentlyContinue)
    if ($keys.Count -eq 0) {
        Write-Log 'No per-user Outlook add-ins were found.' 'WARN'
        return
    }

    $keys | ForEach-Object {
        $item = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
        [pscustomobject]@{ AddIn = $_.PSChildName; LoadBehavior = $item.LoadBehavior; FriendlyName = $item.FriendlyName }
    } | Format-Table -AutoSize

    if (-not (Confirm-Repair 'Disable all per-user Outlook COM add-ins for troubleshooting?' -HighImpact)) { return }
    Stop-OutlookProcess

    $registryBackup = Join-Path $BackupRoot 'Outlook_User_Addins.reg'
    if ($DryRun) {
        Write-Log "Would export the add-in registry key to $registryBackup and set LoadBehavior to 2." 'DRYRUN'
        return
    }

    & reg.exe export 'HKCU\Software\Microsoft\Office\Outlook\Addins' $registryBackup /y | Out-Null
    foreach ($key in $keys) {
        New-ItemProperty -LiteralPath $key.PSPath -Name LoadBehavior -PropertyType DWord -Value 2 -Force | Out-Null
        Write-Log "Disabled user add-in: $($key.PSChildName)" 'SUCCESS'
    }
    Write-Log "Registry backup saved to $registryBackup" 'SUCCESS'
}

function Invoke-RestartSearchService {
    if (-not (Test-IsAdministrator)) {
        Write-Log 'Restarting Windows Search requires Run as administrator.' 'WARN'
        return
    }
    if (-not (Confirm-Repair 'Restart the Windows Search service?')) { return }
    if ($DryRun) {
        Write-Log 'Would restart the WSearch service.' 'DRYRUN'
        return
    }
    Restart-Service WSearch -Force
    Write-Log 'Windows Search service restarted.' 'SUCCESS'
}

function Invoke-RestartClickToRun {
    if (-not (Test-IsAdministrator)) {
        Write-Log 'Restarting Office Click-to-Run requires Run as administrator.' 'WARN'
        return
    }
    if (-not (Confirm-Repair 'Restart the Office Click-to-Run service?')) { return }
    if ($DryRun) {
        Write-Log 'Would restart the ClickToRunSvc service.' 'DRYRUN'
        return
    }
    Restart-Service ClickToRunSvc -Force
    Write-Log 'Office Click-to-Run service restarted.' 'SUCCESS'
}

function Invoke-FlushDns {
    if (-not (Confirm-Repair 'Flush the local DNS resolver cache?')) { return }
    if ($DryRun) {
        Write-Log 'Would run ipconfig.exe /flushdns.' 'DRYRUN'
        return
    }
    & ipconfig.exe /flushdns | Out-Null
    Write-Log 'DNS resolver cache flushed.' 'SUCCESS'
}

function Invoke-RestartOutlook {
    $outlookPath = Get-OutlookPath
    if (-not $outlookPath) {
        Write-Log 'OUTLOOK.EXE was not found.' 'ERROR'
        return
    }
    Stop-OutlookProcess
    if ($DryRun) {
        Write-Log "Would restart $outlookPath." 'DRYRUN'
        return
    }
    Start-Process -FilePath $outlookPath
    Write-Log 'Outlook restarted.' 'SUCCESS'
}

function Invoke-OpenProfileManager {
    if ($DryRun) {
        Write-Log 'Would open the Mail profile manager.' 'DRYRUN'
        return
    }
    Start-Process control.exe -ArgumentList 'mlcfg32.cpl'
    Write-Log 'Opened the Mail profile manager.' 'SUCCESS'
}

function Invoke-OpenOfficeRepair {
    if ($DryRun) {
        Write-Log 'Would open Programs and Features for Microsoft 365 Quick Repair or Online Repair.' 'DRYRUN'
        return
    }
    Start-Process appwiz.cpl
    Write-Log 'Opened Programs and Features. Select Microsoft 365, Change, then Quick Repair or Online Repair.' 'SUCCESS'
}

function Invoke-AllSafeRepairs {
    Write-Log 'Starting the safe Outlook repair workflow.'
    Invoke-RebuildRoamCache
    Invoke-ResetNavigationPane
    Invoke-FlushDns
    if (Test-IsAdministrator) {
        Invoke-RestartSearchService
        Invoke-RestartClickToRun
    }
    Invoke-RestartOutlook
}

function Show-Menu {
    do {
        Clear-Host
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host '  MICROSOFT OUTLOOK REPAIR TOOLKIT' -ForegroundColor Cyan
        Write-Host "  Version $ScriptVersion | Dewald Pretorius" -ForegroundColor DarkCyan
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host "Log: $LogFile"
        Write-Host "Dry run: $DryRun"
        Write-Host
        Write-Host ' 1. Run safe Outlook repairs'
        Write-Host ' 2. Rebuild Outlook RoamCache'
        Write-Host ' 3. Reset navigation pane'
        Write-Host ' 4. Reset all custom views'
        Write-Host ' 5. Rebuild OST cache files (backed up)'
        Write-Host ' 6. Disable per-user Outlook add-ins (registry backup)'
        Write-Host ' 7. Restart Windows Search service'
        Write-Host ' 8. Restart Office Click-to-Run service'
        Write-Host ' 9. Flush DNS cache'
        Write-Host '10. Restart Outlook'
        Write-Host '11. Open Mail profile manager'
        Write-Host '12. Open Microsoft 365 repair'
        Write-Host ' 0. Exit'
        $choice = Read-Host 'Select an option'
        try {
            switch ($choice) {
                '1'  { Invoke-AllSafeRepairs }
                '2'  { Invoke-RebuildRoamCache }
                '3'  { Invoke-ResetNavigationPane }
                '4'  { Invoke-ResetViews }
                '5'  { Invoke-RebuildOstCache }
                '6'  { Invoke-DisableUserAddIns }
                '7'  { Invoke-RestartSearchService }
                '8'  { Invoke-RestartClickToRun }
                '9'  { Invoke-FlushDns }
                '10' { Invoke-RestartOutlook }
                '11' { Invoke-OpenProfileManager }
                '12' { Invoke-OpenOfficeRepair }
                '0'  { return }
                default { Write-Host 'Invalid selection.' -ForegroundColor Yellow }
            }
        } catch {
            Write-Log $_.Exception.Message 'ERROR'
        }
        if ($choice -ne '0') {
            Write-Host
            [void](Read-Host 'Press Enter to continue')
        }
    } while ($true)
}

Write-Log "Outlook Repair Toolkit $ScriptVersion started. DryRun=$DryRun"
Get-OutlookSnapshot -Stage 'Before'

$repairSwitches = @(
    $RepairAllSafe, $RestartOutlook, $RebuildRoamCache, $ResetNavigationPane,
    $ResetViews, $RebuildOstCache, $DisableUserAddIns, $RestartSearchService,
    $RestartClickToRun, $FlushDns, $OpenProfileManager, $OpenOfficeRepair
)

try {
    if (-not ($repairSwitches -contains $true)) {
        Show-Menu
    } else {
        if ($RepairAllSafe)         { Invoke-AllSafeRepairs }
        if ($RebuildRoamCache)      { Invoke-RebuildRoamCache }
        if ($ResetNavigationPane)   { Invoke-ResetNavigationPane }
        if ($ResetViews)            { Invoke-ResetViews }
        if ($RebuildOstCache)       { Invoke-RebuildOstCache }
        if ($DisableUserAddIns)     { Invoke-DisableUserAddIns }
        if ($RestartSearchService)  { Invoke-RestartSearchService }
        if ($RestartClickToRun)     { Invoke-RestartClickToRun }
        if ($FlushDns)              { Invoke-FlushDns }
        if ($RestartOutlook)        { Invoke-RestartOutlook }
        if ($OpenProfileManager)    { Invoke-OpenProfileManager }
        if ($OpenOfficeRepair)      { Invoke-OpenOfficeRepair }
    }
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    $global:LASTEXITCODE = 1
} finally {
    Get-OutlookSnapshot -Stage 'After'
    Write-Log "Repair workflow finished. Backups: $BackupRoot" 'SUCCESS'
    Write-Host "Logs and backups: $OutputPath" -ForegroundColor Green
}
