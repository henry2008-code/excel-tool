@echo off
chcp 65001 >nul
echo ========================================
echo   加密Excel合并工具 - Windows 打包脚本
echo ========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] 检查并安装 PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
)
echo ✓ PyInstaller 已就绪
echo.

echo [2/4] 检查并安装项目依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo [警告] 部分依赖安装失败，继续打包...
)
echo.

echo [3/4] 开始打包...
echo 这可能需要几分钟时间，请耐心等待...
echo.

REM 清理旧的打包文件
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"

REM 执行打包
pyinstaller --name "加密Excel合并工具" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data "requirements.txt;." ^
    --hidden-import openpyxl ^
    --hidden-import olefile ^
    --hidden-import cryptography ^
    --hidden-import msoffcrypto ^
    --hidden-import PyQt6 ^
    --hidden-import PyQt6.sip ^
    --exclude-module tkinter ^
    --exclude-module unittest ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo [错误] 打包失败，请检查上方错误信息
    pause
    exit /b 1
)

echo.
echo [4/4] 打包完成！
echo.
echo ========================================
echo   输出文件位置: dist\加密Excel合并工具.exe
echo ========================================
echo.

if exist "dist\加密Excel合并工具.exe" (
    echo ✓ EXE 文件已生成
    echo.
    echo 文件大小:
    for %%A in ("dist\加密Excel合并工具.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1048576
        echo   %%~zA 字节 (!sizeMB! MB)
    )
    echo.
    echo 提示: 该 exe 文件可在未安装 Python 的 Windows 电脑上直接运行
    echo.
    pause
    
    REM 询问是否打开输出目录
    set /p open_dir="是否打开输出目录? (Y/N): "
    if /i "%open_dir%"=="Y" (
        explorer "dist"
    )
) else (
    echo [错误] 未找到生成的 EXE 文件
    pause
    exit /b 1
)
