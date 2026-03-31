@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title DataEngine 打包程序
echo ========================================
echo   [打包任务] 正在为 DataEngine 生成 exe...
echo ========================================

:: 执行打包命令
pyinstaller --noconsole ^
--onefile ^
--collect-all customtkinter ^
--collect-all selenium ^
--add-data "resources;resources" ^
--icon="app_icon.ico" ^
--name "DataEngine" ^
GUI.py



if %ERRORLEVEL% EQU 0 (
    echo.
    echo [成功] 打包完成！请在 'dist' 文件夹下查看。
) else (
    echo.
    echo [错误] 打包失败，请检查上面的错误日志。
)

echo ========================================
pause