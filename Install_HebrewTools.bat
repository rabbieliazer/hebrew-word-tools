@echo off
setlocal enabledelayedexpansion
title Tosh Developers - Hebrew Tools Installer

:: --- CONFIGURATION (UPDATED TO RAW LINKS) ---
set "REPO_URL=https://raw.githubusercontent.com/rabbieliazer/hebrew-word-tools/main"
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
:: We use -L to follow redirects and -k to skip SSL certificate issues if they occur
for /f "delims=" %%a in ('curl -s -L -k %REPO_URL%/version.txt') do set "LATEST_VER=%%a"

:: Remove any potential trailing spaces or hidden characters
set "LATEST_VER=%LATEST_VER: =%"

if "%LATEST_VER%"=="" (
    echo [ERROR] Could not fetch version information. 
    echo Please check your internet or the GitHub repository status.
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

curl -s -L -k -o "%STARTUP_DIR%\%DOTM_NAME%" "%REPO_URL%/%DOTM_NAME%"

if %errorlevel% equ 0 (
    :: This removes the "Security Risk" block from Windows
    powershell -command "Unblock-File -Path '%STARTUP_DIR%\%DOTM_NAME%'"
    echo %LATEST_VER% > "%VERSION_FILE%"
    echo.
    echo [SUCCESS] Hebrew Tools v%LATEST_VER% installed successfully!
    echo Restart Word to see the changes.
) else (
    echo [ERROR] Download failed.
)

pause
