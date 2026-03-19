@echo off
REM 养老金融通报数据自动化处理系统
REM 打包脚本 (Windows)
REM
REM 使用方法: build.bat

echo ========================================
echo 养老金融通报数据处理系统 - 打包脚本
echo ========================================
echo.

REM 检查Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo 错误: 未找到Python
    pause
    exit /b 1
)

REM 安装依赖
echo 检查依赖...
pip install -r requirements.txt >nul 2>&1

REM 安装PyInstaller
pip install pyinstaller >nul 2>&1

REM 创建输出目录
if not exist "dist" mkdir dist

REM 打包
echo.
echo 开始打包...
echo.

pyinstaller pension_financial.spec --clean

echo.
echo ========================================
echo 打包完成！
echo 输出目录: dist\养老金融通报数据处理系统
echo ========================================
echo.
pause
