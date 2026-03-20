@echo off
chcp 65001 >nul
title DataEngine 打包程序
echo ========================================
echo   [打包任务] 正在为 DataEngine 生成 exe...
echo ========================================

:: 进入根目录执行，防止路径混乱
cd ..

:: 执行打包命令
".venv\Scripts\pyinstaller" --noconsole ^
--onefile ^
--collect-all customtkinter ^
--icon="resources\app_icon.ico" ^
--name "DataEngine" ^
--add-data "config.yaml;." ^
--add-data "resources;resources" ^
gui.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [成功] 打包完成！请在 'dist' 文件夹下查看。
) else (
    echo.
    echo [错误] 打包失败，请检查上面的错误日志。
)

echo ========================================
pause