@echo off
title AutoGeneratePDF 打包程序
color 0a

:: 启用 UTF-8 编码，防止中文乱码
chcp 65001 > nul

echo.
echo =========================================================
echo     AutoGeneratePDF v3.1 - 一键打包为 .exe
echo =========================================================
echo.
echo     此脚本会将 print_tool_gui.py 和 msedgedriver.exe
echo     打包成一个独立的 exe 文件。
echo.

:: 检查依赖文件是否存在
if not exist "print_tool_gui.py" (
    echo ❌ 错误: 核心脚本 'print_tool_gui.py' 不存在！
    pause
    exit /b
)
if not exist "msedgedriver.exe" (
    echo ❌ 错误: 驱动文件 'msedgedriver.exe' 不存在！
    echo    请确保驱动和此脚本在同一目录下。
    pause
    exit /b
)

:: (可选) 检查 requirements.txt 并安装依赖
if exist "requirements.txt" (
    echo ✅ 正在检查并安装依赖库...
    pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo ❌ 依赖安装失败，请检查 Python 和 pip 环境。
        pause
        exit /b
    )
)

:: 检查 pyinstaller 是否安装
where pyinstaller >nul
if errorlevel 1 (
    echo 🟡 PyInstaller 未安装，正在自动安装...
    pip install pyinstaller --quiet
)

:: 清理旧的打包文件
echo 🧹 正在清理旧的打包文件 (dist, build, .spec)...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist print_tool_gui.spec del print_tool_gui.spec

echo.
echo 🚀 正在打包为单文件 exe，请稍等片刻...
echo.

:: --- 核心打包命令 ---
:: --onefile: 打包成单个 exe 文件
:: --windowed: 运行时不显示黑色命令行窗口
:: --icon="icon.ico": 为 exe 设置图标 (确保 icon.ico 存在)
:: --add-data "msedgedriver.exe;.":  <-- 这是最关键的一步！
::     它告诉 PyInstaller:
::     1. 将 "msedgedriver.exe" 文件添加进来。
::     2. 在程序运行时，将它解压到程序的根目录 (用“.”表示)。
pyinstaller ^
    --onefile ^
    --windowed ^
    --icon="icon.ico" ^
    --add-data "msedgedriver.exe;." ^
    print_tool_gui.py

if errorlevel 1 (
    echo.
    echo ❌ 打包失败！请检查上方的错误信息。
    pause
    exit /b
)

echo.
echo =========================================================
echo ✨ 打包成功！✨
echo =========================================================
echo.
echo 🔹 输出位置：dist\print_tool_gui.exe
echo.
echo ➡️  现在您可以将 dist 文件夹中的 print_tool_gui.exe
echo     文件发送给任何 Windows 用户直接运行！
echo.
pause