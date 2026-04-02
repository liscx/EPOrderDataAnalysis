@echo off
title DataEngine Cleaner
echo ========================================
echo   [TASK] Removing build artifacts...
echo ========================================

:: Move to root level from scripts folder
cd /d %~dp0\..

if exist build (
    rd /s /q build
    echo Removed: build/
)
if exist dist (
    rd /s /q dist
    echo Removed: dist/
)
if exist DataEngine.spec (
    del /q DataEngine.spec
    echo Removed: DataEngine.spec
)

echo.
echo   [DONE] Clean finished!
echo ========================================
pause