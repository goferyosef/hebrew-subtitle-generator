@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title Hebrew Subtitle Generator - Setup

:: ── Header ───────────────────────────────────────────────────────────────────
powershell -NoProfile -Command ^
  "Write-Host ''; Write-Host '  Hebrew Subtitle Generator' -ForegroundColor Cyan -NoNewline; Write-Host ' - Setup' -ForegroundColor White; Write-Host '  ================================================' -ForegroundColor DarkCyan; Write-Host ''"

:: ── Find Python ──────────────────────────────────────────────────────────────
set "PYTHON_EXE="

:: 1. py launcher (most reliable on Windows)
py -3 --version >nul 2>&1
if !errorlevel! == 0 ( set "PYTHON_EXE=py -3" & goto :HavePython )

:: 2. python in PATH
python --version >nul 2>&1
if !errorlevel! == 0 ( set "PYTHON_EXE=python" & goto :HavePython )

:: 3. python3 in PATH
python3 --version >nul 2>&1
if !errorlevel! == 0 ( set "PYTHON_EXE=python3" & goto :HavePython )

:: 4. Common per-user install locations (Windows 10/11 default)
for %%V in (313 312 311 310) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto :HavePython
    )
)

:: 5. Common system-wide locations
for %%V in (313 312 311 310) do (
    if exist "C:\Python%%V\python.exe" (
        set "PYTHON_EXE=C:\Python%%V\python.exe"
        goto :HavePython
    )
)

:: ── Python not found — try winget ────────────────────────────────────────────
powershell -NoProfile -Command "Write-Host '  Python not found.' -ForegroundColor Yellow"

winget --version >nul 2>&1
if !errorlevel! == 0 (
    powershell -NoProfile -Command "Write-Host '  Installing Python 3.12 via winget...' -ForegroundColor Cyan"
    winget install Python.Python.3.12 ^
        --accept-package-agreements --accept-source-agreements --silent
    if !errorlevel! == 0 (
        :: Refresh PATH from registry
        for /f "usebackq tokens=2*" %%a in (
            `reg query "HKCU\Environment" /v Path 2^>nul`
        ) do set "HKCU_PATH=%%b"
        for /f "usebackq tokens=2*" %%a in (
            `reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul`
        ) do set "HKLM_PATH=%%b"
        set "PATH=!HKLM_PATH!;!HKCU_PATH!;!PATH!"

        py -3 --version >nul 2>&1
        if !errorlevel! == 0 ( set "PYTHON_EXE=py -3"  & goto :HavePython )
        python --version >nul 2>&1
        if !errorlevel! == 0 ( set "PYTHON_EXE=python" & goto :HavePython )
    )
)

:: ── Still no Python ───────────────────────────────────────────────────────────
echo.
powershell -NoProfile -Command ^
  "Write-Host '  ERROR: Python could not be installed automatically.' -ForegroundColor Red; ^
   Write-Host ''; ^
   Write-Host '  Please install Python 3.10 or newer from:' -ForegroundColor Yellow; ^
   Write-Host '    https://www.python.org/downloads/' -ForegroundColor White; ^
   Write-Host '  IMPORTANT: tick  Add Python to PATH  during installation!' -ForegroundColor Yellow; ^
   Write-Host '  Then run INSTALL.bat again.' -ForegroundColor White; ^
   Write-Host ''"
pause
exit /b 1

:HavePython
:: ── Launch the GUI installer ──────────────────────────────────────────────────
powershell -NoProfile -Command "Write-Host '  Python found. Launching installer...' -ForegroundColor Green; Write-Host ''"
%PYTHON_EXE% installer.py
if !errorlevel! neq 0 (
    echo.
    powershell -NoProfile -Command "Write-Host '  Installer exited with an error.' -ForegroundColor Red"
    pause
)
exit /b 0
