@echo off
chcp 65001 >nul
echo ========================================
echo   [清理任务] 正在移除构建缓存...
echo ========================================

:: 进入根目录执行
cd ..

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
if exist "build_tools\DataEngine.spec" (
    del /q "build_tools\DataEngine.spec"
    echo 已移除: DataEngine.spec
)
if exist "DataEngine.spec" (
    del /q "DataEngine.spec"
    echo 已移除: DataEngine.spec
)

echo.
echo   清理完成！
echo ========================================
pause