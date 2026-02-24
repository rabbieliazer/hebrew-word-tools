@echo off
setlocal enabledelayedexpansion
title Hebrew Tools Installer

set "DOTM_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/HebrewTools.dotm"
set "VERSION_URL=https://github.com/rabbieliazer/hebrew-word-tools/raw/refs/heads/main/version.txt"
set "TARGET_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "INTERNAL_DIR=%APPDATA%\Microsoft\Word"
set "FILE_NAME=HebrewTools.dotm"
set "LOCAL_VER_FILE=%INTERNAL_DIR%\ht_version.txt"
set "PERM_BAT=%INTERNAL_DIR%\update_script.bat"

echo ==========================================
echo       HEBREW TOOLS - VERSION 4.1
echo ==========================================

copy /y "%~f0" "%PERM_BAT%" >nul

tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [WARNING] Word is open. CLOSE IT NOW to update correctly.
    pause
)

for /f "delims=" %%a in ('curl -k -s -L %VERSION_URL%') do set "LATEST_VER=%%a"
if "%LATEST_VER%"=="" (echo GitHub connection error. && pause && exit)

if exist "%LOCAL_VER_FILE%" (set /p CURRENT_VER=<"%LOCAL_VER_FILE%") else (set "CURRENT_VER=0")

if "%LATEST_VER%"=="%CURRENT_VER%" (
    if exist "%TARGET_DIR%\%FILE_NAME%" (
        echo [INFO] Already updated to v%LATEST_VER%.
        timeout /t 3
        exit
    )
)

echo Downloading v%LATEST_VER%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
curl -k -L -o "%TARGET_DIR%\%FILE_NAME%" "%DOTM_URL%"

if %errorlevel% equ 0 (
    powershell -command "Unblock-File -Path '%TARGET_DIR%\%FILE_NAME%'"
    echo %LATEST_VER% > "%LOCAL_VER_FILE%"
    echo [SUCCESS] Install Complete! Restart Word.
) else (
    echo [ERROR] Download failed.
)
pause
