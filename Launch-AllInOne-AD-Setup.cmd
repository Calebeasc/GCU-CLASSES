@echo off
setlocal
set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%AllInOne-AD-Setup.ps1
if not exist "%SCRIPT_PATH%" (
  echo Could not locate AllInOne-AD-Setup.ps1 next to this launcher.
  pause
  exit /b 1
)
powershell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -File "%SCRIPT_PATH%"
endlocal
