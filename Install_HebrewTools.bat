@echo off
title Tosh Developers - Emergency Installer
echo --------------------------------------------------
echo CLEANING OLD FILES...
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "URL=https://raw.githubusercontent.com/rabbieliazer/hebrew-word-tools/main/HebrewTools.dotm"

:: Close Word if open
taskkill /f /im winword.exe >nul 2>&1

:: Create directory if missing
if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%"

echo.
echo DOWNLOADING FROM GITHUB...
curl -L -k -o "%STARTUP_DIR%\HebrewTools.dotm" "%URL%"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Download failed. Please check your internet connection.
    pause
    exit
)

echo.
echo UNBLOCKING MACROS...
powershell -Command "Unblock-File -Path '%STARTUP_DIR%\HebrewTools.dotm'"

echo.
echo --------------------------------------------------
echo SUCCESS! Hebrew Tools has been installed.
echo You can now open Microsoft Word.
echo --------------------------------------------------
pause
