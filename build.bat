@echo off
echo ===================================================
echo   正在开始打包 Data Engine v6.2 ...
echo ===================================================

:: 1. 清理旧的打包文件（防止缓存干扰）
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist DataEngine.spec del /f /q DataEngine.spec

:: 2. 执行 PyInstaller 命令
:: 注意：确保你的图标文件名是 de_dark_icon.ico
pyinstaller --noconsole ^
            --onefile ^
            --collect-all customtkinter ^
            --icon="de_dark_icon.ico" ^
            --name "DataEngine" ^
            GUI.py

echo ===================================================
echo   打包完成！请在 dist 文件夹中查看 DataEngine.exe
echo ===================================================
pause