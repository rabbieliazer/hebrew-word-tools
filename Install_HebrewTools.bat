@echo off
setlocal enabledelayedexpansion
title Tosh Developers - Hebrew Tools Installer

:: --- CONFIGURATION (CORRECTED RAW LINKS) ---
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
echo Checking for updates at %REPO_URL%/version.txt...
for /f "delims=" %%a in ('curl -s -L -k %REPO_URL%/version.txt') do set "LATEST_VER=%%a"

:: Clean up any potential hidden characters
set "LATEST_VER=%LATEST_VER: =%"

if "%LATEST_VER%"=="" (
    echo [ERROR] Could not connect to GitHub or file is empty.
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

:: Using -L to follow redirects and -k for SSL compatibility
curl -k -L -o "%STARTUP_DIR%\%DOTM_NAME%" "%REPO_URL%/%DOTM_NAME%"

if %errorlevel% equ 0 (
    powershell -command "Unblock-File -Path '%STARTUP_DIR%\%DOTM_NAME%'"
    echo %LATEST_VER% > "%VERSION_FILE%"
    echo.
    echo [SUCCESS] Hebrew Tools v%LATEST_VER% installed successfully!
    echo Restart Word to see the changes.
) else (
    echo [ERROR] Download failed.
)

pause
