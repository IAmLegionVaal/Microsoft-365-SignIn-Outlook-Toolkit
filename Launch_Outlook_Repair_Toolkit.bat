@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%~dp0Outlook_Repair_Toolkit.ps1' -ErrorAction SilentlyContinue; & '%~dp0Outlook_Repair_Toolkit.ps1'"
echo.
echo Outlook repair workflow finished.
pause
endlocal
