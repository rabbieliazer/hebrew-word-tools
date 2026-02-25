@echo off
setlocal enabledelayedexpansion
title Tosh Developers - Hebrew Tools Installer

:: --- CONFIGURATION ---
set "REPO_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main"
set "DOTM_NAME=HebrewTools.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "VERSION_FILE=%APPDATA%\Microsoft\Word\ht_version.txt"

echo ==================================================
echo           TOSH DEVELOPERS - HEBREW TOOLS
echo ==================================================
echo.

:: 1. CHECK IF WORD IS OPEN
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [ERROR] Microsoft Word is currently running.
    echo Please CLOSE Word and then press any key to continue.
    pause >nul
)

:: 2. GET LATEST VERSION FROM GITHUB
echo Checking for updates...
for /f "delims=" %%a in ('curl -s -L %REPO_URL%/version.txt') do set "LATEST_VER=%%a"

if "%LATEST_VER%"=="" (
    echo [ERROR] Could not connect to GitHub. Check your internet.
    pause
    exit
)

:: 3. COMPARE VERSIONS
if exist "%VERSION_FILE%" (
    set /p CURRENT_VER=<"%VERSION_FILE%"
) else (
    set "CURRENT_VER=0"
)

if "%LATEST_VER%"=="%CURRENT_VER%" (
    echo [INFO] You are already on the latest version (v%LATEST_VER%).
    timeout /t 3
    exit
)

:: 4. DOWNLOAD AND UNBLOCK
echo New version found: v%LATEST_VER%
echo Downloading %DOTM_NAME%...

if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%"

curl -k -L -o "%STARTUP_DIR%\%DOTM_NAME%" "%REPO_URL%/%DOTM_NAME%"

if %errorlevel% equ 0 (
    :: Remove "Mark of the Web" so the ribbon actually shows up
    powershell -command "Unblock-File -Path '%STARTUP_DIR%\%DOTM_NAME%'"
    echo %LATEST_VER% > "%VERSION_FILE%"
    echo.
    echo [SUCCESS] Hebrew Tools v%LATEST_VER% installed successfully!
    echo You can now open Microsoft Word.
) else (
    echo [ERROR] Download failed. Please try again later.
)

pause
