@echo off
cd /d "%~dp0"

if not exist "runtime\python\pythonw.exe" (
  echo [ERROR] runtime\python not found. Please unzip the whole folder.
  pause
  exit /b 1
)
if not exist "runtime\node\tyc.cmd" (
  echo [ERROR] runtime\node not found. Please unzip the whole folder.
  pause
  exit /b 1
)

set "PATH=%~dp0runtime\node;%PATH%"
set "TCL_LIBRARY=%~dp0runtime\python\tcl\tcl8.6"
set "TK_LIBRARY=%~dp0runtime\python\tcl\tk8.6"

start "" /D "%~dp0" "%~dp0runtime\python\pythonw.exe" app.py
