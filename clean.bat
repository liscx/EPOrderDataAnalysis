@echo off
echo ========================================
echo   [清理任务] 正在移除构建缓存...
echo ========================================

:: 删除 PyInstaller 生成的文件夹
if exist build (
    rd /s /q build
    echo 已移除: build 文件夹
)
if exist dist (
    rd /s /q dist
    echo 已移除: dist 文件夹
)

:: 删除生成的 spec 配置文件
if exist "DataEngine.spec" (
    del /q "DataEngine.spec"
    echo 已移除: DataEngine.spec
)

echo.
echo   清理完成！
echo ========================================
pause