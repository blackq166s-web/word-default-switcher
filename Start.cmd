@echo off
setlocal
if not exist "%~dp0app\Window.ps1" (
  echo Missing app files. Please keep Start.cmd together with the app folder.
  pause
  exit /b 1
)
attrib +h "%~dp0app" >nul 2>&1
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0app\Window.ps1" %*
if errorlevel 1 pause
