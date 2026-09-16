@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0MD_Table_Converter.ps1"
if errorlevel 1 (
  echo.
  echo Converter failed to start.
  echo Press any key to close this window.
  pause >nul
)
