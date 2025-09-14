@echo off
title 打印工具打包程序
color 0a

:: 启用 UTF-8 编码，防止中文乱码
chcp 65001 > nul

echo.
echo ========================================
echo     打印工具 - 一键打包为 .exe (桌面保存版)
echo ========================================
echo.

:: 安装依赖
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ❌ 依赖安装失败，请确保已安装 Python 和 pip
    pause
    exit /b
)

:: 检查 pyinstaller 是否可用
where pyinstaller >nul
if errorlevel 1 (
    echo ❌ PyInstaller 未安装，正在安装...
    pip install pyinstaller --quiet
)

:: 清理旧的 dist 和 build 文件夹
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

echo.
echo ✅ 正在打包为单文件 exe，请稍等...

:: 👇 关键：不再添加任何外部文件，由 webdriver-manager 自动处理
pyinstaller ^
    --onefile ^
    --windowed ^
    print_tool_gui.py

if errorlevel 1 (
    echo ❌ 打包失败，请查看上方错误信息
    pause
    exit /b
)

echo.
echo ✅ 打包成功！
echo.
echo 🔹 输出位置：dist\print_tool_gui.exe
echo.
echo 🚀 可直接发送给用户使用！
echo.
pause