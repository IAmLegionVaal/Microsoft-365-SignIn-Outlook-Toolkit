#requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 Sign-In Outlook Toolkit.
.DESCRIPTION
    Diagnostic-only Microsoft 365 app installation and connectivity checker.
#>
[CmdletBinding()]
param([string]$OutputPath)

$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'M365_Outlook_Reports' }
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
function New-Check { param($Category,$Name,$Status,$Value,$Recommendation) [PSCustomObject]@{Category=$Category;Name=$Name;Status=$Status;Value=$Value;Recommendation=$Recommendation} }
$checks = @()

$officePaths = @("$env:ProgramFiles\Microsoft Office\root\Office16","${env:ProgramFiles(x86)}\Microsoft Office\root\Office16")
foreach($app in @('OUTLOOK.EXE','WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE')){
    $found = $officePaths | ForEach-Object { Join-Path $_ $app } | Where-Object { Test-Path $_ } | Select-Object -First 1
    $checks += New-Check 'Office Apps' $app ($(if($found){'OK'}else{'Info'})) $found 'Confirm expected Microsoft 365 app is installed.'
}
try { $svc = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue; $checks += New-Check 'Office Service' 'ClickToRun service' ($(if($svc -and $svc.Status -eq 'Running'){'OK'}else{'Info'})) "Status: $($svc.Status)" 'Required for Microsoft 365 Apps servicing.' } catch {}
foreach($processName in @('OUTLOOK','OneDrive','Teams')){
    $proc = Get-Process $processName -ErrorAction SilentlyContinue
    $checks += New-Check 'Processes' $processName 'Info' (@($proc).Count) 'Process count for support context.'
}
foreach($hostName in @('login.microsoftonline.com','outlook.office.com','autodiscover-s.outlook.com','officecdn.microsoft.com')){
    try { [void][System.Net.Dns]::GetHostAddresses($hostName); $dns='Resolved' } catch { $dns='DNS failed' }
    try { $tcp = Test-NetConnection -ComputerName $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue } catch { $tcp=$false }
    $checks += New-Check 'Connectivity' $hostName ($(if($tcp){'OK'}else{'Warning'})) "DNS: $dns; TCP443: $tcp" 'Check DNS, firewall, proxy, and internet access if this fails.'
}
$csv = Join-Path $OutputPath "m365_outlook_checks_$RunStamp.csv"
$json = Join-Path $OutputPath "m365_outlook_checks_$RunStamp.json"
$html = Join-Path $OutputPath "m365_outlook_report_$RunStamp.html"
$checks | Export-Csv $csv -NoTypeInformation -Encoding UTF8
$checks | ConvertTo-Json -Depth 5 | Set-Content $json -Encoding UTF8
$checks | ConvertTo-Html -Title 'M365 Outlook Diagnostic' -PreContent "<h1>Microsoft 365 Outlook Diagnostic - $env:COMPUTERNAME</h1><p>Generated $(Get-Date)</p>" | Set-Content $html -Encoding UTF8
$checks | Format-Table -AutoSize -Wrap
Write-Host "Reports saved to: $OutputPath" -ForegroundColor Green
Start-Process explorer.exe -ArgumentList "`"$OutputPath`"" -ErrorAction SilentlyContinue
