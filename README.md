# Microsoft 365 Sign-In and Outlook Toolkit

A PowerShell toolkit for diagnosing and repairing common classic Microsoft Outlook and Microsoft 365 sign-in problems.

## Included tools

### Diagnostic toolkit

`M365_SignIn_Outlook_Toolkit.ps1`

Checks:

- Office Click-to-Run installation and service state
- Installed Microsoft 365 application paths
- Outlook profile registry information
- OST and PST context
- Microsoft 365 endpoint DNS and TCP 443 connectivity
- Autodiscover context
- Proxy configuration
- OneDrive and Microsoft 365 process context
- HTML, CSV and JSON reports

Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\M365_SignIn_Outlook_Toolkit.ps1
```

### Outlook repair toolkit

`Outlook_Repair_Toolkit.ps1`

Repairs:

- Outlook startup and crash problems caused by corrupted RoamCache data
- Broken or missing navigation panes
- Corrupted custom folder views
- Damaged OST cache files by moving them to a timestamped backup before Outlook recreates them
- Problematic per-user COM add-ins by exporting their registry configuration and disabling them
- Outlook search problems by restarting Windows Search
- Office servicing problems by restarting Click-to-Run
- Connection problems by flushing the DNS resolver cache
- Profile problems through the Mail profile manager
- Microsoft 365 installation problems through Quick Repair or Online Repair

The script creates before-and-after JSON snapshots, timestamped logs and recoverable backups. PST files are never changed.

Run the menu:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Outlook_Repair_Toolkit.ps1
```

Preview without changing anything:

```powershell
.\Outlook_Repair_Toolkit.ps1 -RepairAllSafe -DryRun
```

Run selected repairs:

```powershell
.\Outlook_Repair_Toolkit.ps1 -RebuildRoamCache -ResetNavigationPane -FlushDns
.\Outlook_Repair_Toolkit.ps1 -RebuildOstCache
.\Outlook_Repair_Toolkit.ps1 -DisableUserAddIns
```

A double-click launcher is also included:

```text
Launch_Outlook_Repair_Toolkit.bat
```

## Safety

- RoamCache and OST files are moved to a timestamped backup folder instead of being deleted.
- PST files are never modified.
- Add-in registry settings are exported before changes are made.
- High-impact repairs require the technician to type `REPAIR`.
- Standard repairs require explicit confirmation.
- `-DryRun` previews the workflow without making changes.
- Administrative repairs are skipped with a warning when the script is not elevated.

## Validation status

This toolkit has been tested successfully by the author on his own Windows machines using classic Outlook and Microsoft 365. The documented diagnostic, backup and repair workflows completed as intended in those environments.

Results may vary with the Windows and Office build, Outlook profile type, Exchange or Microsoft 365 tenant configuration, add-ins, permissions, security policy, network conditions and user-specific profile data. Successful author testing does not guarantee identical behaviour in every environment, so use `-DryRun` and validate on a non-critical profile when introducing the toolkit to a new configuration.

## Author

Dewald Pretorius — L2 IT Support Engineer
