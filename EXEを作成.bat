@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo   Excel data tool  -  build EXE
echo ============================================
echo.

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )

if not defined PY (
  echo Python not found. Trying to install via winget...
  where winget >nul 2>&1
  if errorlevel 1 (
    echo [!] winget is not available.
    echo     Please install Python from https://www.python.org/downloads/
    echo     ^(check "Add python.exe to PATH"^) then run this file again.
    pause
    exit /b 1
  )
  winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
  echo.
  echo Python installed. Please CLOSE this window and run this file again.
  pause
  exit /b 0
)

echo Using Python: %PY%
%PY% --version
echo.
%PY% build_exe.py
echo.
echo (Log saved to build_log.txt)
pause
