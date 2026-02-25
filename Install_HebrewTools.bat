@echo off
:: This prevents the window from closing so you can read the error
setlocal enabledelayedexpansion
title Tosh Developers Installer

echo Setting up...

:: 1. SET THE URLS
set "REPO_URL=https://raw.githubusercontent.com/rabbieliazer/hebrew-word-tools/main"
set "DOTM_NAME=HebrewTools.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"

echo Checking connection to: %REPO_URL%/version.txt

:: 2. TEST CURL (The most common point of failure)
curl -s -L -k %REPO_URL%/version.txt > nul
if %errorlevel% neq 0 (
    echo [ERROR] CURL failed. Your computer might be blocking the download.
    pause
    exit
)

:: 3. GET THE VERSION
for /f "tokens=*" %%a in ('curl -s -L -k %REPO_URL%/version.txt') do set "LATEST_VER=%%a"

echo Latest Version found: "%LATEST_VER%"

:: 4. DOWNLOAD THE FILE
echo Downloading to: %STARTUP_DIR%
if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%"

curl -L -k -o "%STARTUP_DIR%\%DOTM_NAME%" "%REPO_URL%/%DOTM_NAME%"

:: 5. UNBLOCK (Using a safer PowerShell call)
echo Unblocking file...
powershell.exe -Command "Unblock-File -Path '%STARTUP_DIR%\%DOTM_NAME%'"

echo.
echo ==========================================
echo DONE! Please open Word and check the Ribbon.
echo ==========================================
pause
