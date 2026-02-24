@echo off
setlocal enabledelayedexpansion
title Hebrew Tools - Smart Update & Installer

:: --- SETTINGS ---
set "DOTM_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/HebrewTools.dotm"
set "VERSION_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/version.txt"
set "TARGET_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "FILE_NAME=HebrewTools.dotm"
set "LOCAL_VER_FILE=%TARGET_DIR%\ht_version.txt"
:: ----------------

echo ==========================================
echo       HEBREW TOOLS - LIVE INSTALLER
echo ==========================================
echo.

:: 1. Check if Word is running
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [WARNING] Microsoft Word is currently running.
    echo Please save your work and CLOSE Word to continue.
    echo.
    pause
    goto :1
)

:1
:: 2. Get the latest version number from GitHub
echo Checking for updates...
for /f "delims=" %%a in ('curl -s -L %VERSION_URL%') do set "LATEST_VER=%%a"

:: 3. Check what version is currently installed
if exist "%LOCAL_VER_FILE%" (
    set /p CURRENT_VER=<"%LOCAL_VER_FILE%"
) else (
    set "CURRENT_VER=0"
)

echo Current Version: %CURRENT_VER%
echo Latest Version:  %LATEST_VER%

:: 4. Comparison Logic
if "%LATEST_VER%"=="%CURRENT_VER%" (
    if exist "%TARGET_DIR%\%FILE_NAME%" (
        echo.
        echo [INFO] You already have the latest version.
        timeout /t 5
        exit
    )
)

echo.
echo [NEW] Downloading Hebrew Tools v%LATEST_VER%...

:: 5. Create Startup folder if missing
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 6. Download the file
curl -L -o "%TARGET_DIR%\%FILE_NAME%" "%DOTM_URL%"

if %errorlevel% equ 0 (
    echo %LATEST_VER% > "%LOCAL_VER_FILE%"
    echo.
    echo [SUCCESS] Hebrew Tools has been installed/updated!
    echo Location: %TARGET_DIR%
    echo.
    echo You can now open Microsoft Word.
) else (
    echo.
    echo [ERROR] Download failed. Check your internet connection.
)

echo ==========================================
pause