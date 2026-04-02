@echo off
title DataEngine Packager
echo ========================================
echo   [TASK] Building DataEngine.exe...
echo ========================================

:: Move to root level implicitly via relative paths
cd /d %~dp0\..

:: Build command
pyinstaller --noconsole ^
--onefile ^
--collect-all customtkinter ^
--collect-all selenium ^
--add-data "resources;resources" ^
--icon="resources\app_icon.ico" ^
--name "DataEngine" ^
GUI.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCCESS] Package finished! Check 'dist' folder.
) else (
    echo.
    echo [ERROR] Build failed. Check logs.
)

echo ========================================
pause