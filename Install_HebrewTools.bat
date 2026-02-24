@echo off
setlocal enabledelayedexpansion
title Hebrew Tools Installer

:: --- SETTINGS ---
set "DOTM_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/HebrewTools.dotm"
set "VERSION_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/version.txt"
set "TARGET_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "INTERNAL_DIR=%APPDATA%\Microsoft\Word"
set "FILE_NAME=HebrewTools.dotm"
set "LOCAL_VER_FILE=%INTERNAL_DIR%\ht_version.txt"
set "PERM_BAT=%INTERNAL_DIR%\update_script.bat"
:: ----------------

echo ==========================================
echo       HEBREW TOOLS - LIVE INSTALLER
echo ==========================================
echo.

:: 1. Save copy for the Word "Update" button
copy /y "%~f0" "%PERM_BAT%" >nul

:: 2. Check Word
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [WARNING] Microsoft Word is open.
    echo Please CLOSE Word now to finish.
    pause
)

:: 3. Check Version
echo Connecting to GitHub...
for /f "delims=" %%a in ('curl -k -s -L %VERSION_URL%') do set "LATEST_VER=%%a"
if "%LATEST_VER%"=="" (echo Error connecting to GitHub. && pause && exit)

if exist "%LOCAL_VER_FILE%" (set /p CURRENT_VER=<"%LOCAL_VER_FILE%") else (set "CURRENT_VER=0")

echo Current Version: %CURRENT_VER%
echo Latest Version:  %LATEST_VER%

if "%LATEST_VER%"=="%CURRENT_VER%" (
    if exist "%TARGET_DIR%\%FILE_NAME%" (
        echo.
        echo [INFO] You are up to date.
        timeout /t 3
        exit
    )
)

:: 4. Install
echo Downloading Hebrew Tools v%LATEST_VER%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
curl -k -L -o "%TARGET_DIR%\%FILE_NAME%" "%DOTM_URL%"

if %errorlevel% equ 0 (
    powershell -command "Unblock-File -Path '%TARGET_DIR%\%FILE_NAME%'"
    echo %LATEST_VER% > "%LOCAL_VER_FILE%"
    echo.
    echo [SUCCESS] Install Complete! Restart Word.
) else (
    echo.
    echo [ERROR] Download failed.
)
pause
