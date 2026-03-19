@echo off
chcp 65001 >nul
REM ============================================================
REM 养老金融通报数据自动化处理系统
REM Windows 一键打包脚本
REM ============================================================

echo.
echo ============================================================
echo   养老金融通报数据处理系统 - Windows 打包工具
echo ============================================================
echo.

REM 检查Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 获取Python版本
python --version >temp_ver.txt 2>nul
set /p PY_VER=<temp_ver.txt
del temp_ver.txt

REM 过滤版本号
echo %PY_VER% | find "Python 3" >nul
if %errorlevel% neq 0 (
    echo [警告] 检测到Python版本可能不符合要求
    echo 推荐使用 Python 3.10 或更高版本
)

echo [信息] 检测到Python环境: %PY_VER%
echo.

REM 安装依赖
echo ============================================================
echo   步骤1/3: 安装依赖包
echo ============================================================
echo.

echo正在安装依赖包，请稍候...
pip install -r requirements.txt --quiet --disable-pip-version-check

if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装可能存在问题，尝试继续...
)

echo.
echo ============================================================
echo   步骤2/3: 安装PyInstaller
echo ============================================================
echo.

pip install pyinstaller --quiet --disable-pip-version-check

if %errorlevel% neq 0 (
    echo [错误] PyInstaller安装失败
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   步骤3/3: 打包应用程序
echo ============================================================
echo.

REM 创建dist目录
if not exist "dist" mkdir dist

REM 执行打包
echo正在打包，请稍候（首次打包可能需要3-5分钟）...
echo.

pyinstaller main.py --name "养老金融通报数据处理系统" --windowed --onefile ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=numpy ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.scrolledtext ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=core.config ^
    --hidden-import=core.processor ^
    --hidden-import=utils.excel_tool ^
    --hidden-import=utils.init_template ^
    --clean

if %errorlevel% neq 0 (
    echo.
    echo [错误] 打包过程中出现错误
    echo 请检查上方错误信息
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   打包完成！
echo ============================================================
echo.

REM 查找生成的EXE文件
for /r "dist" %%f in (*.exe) do (
    set "EXE_PATH=%%f"
)

if defined EXE_PATH (
    echo [成功] EXE文件已生成:
    echo   %EXE_PATH%
    echo.
    echo 文件大小:
    for %%A in ("%EXE_PATH%") do echo   %%~zA 字节
    echo.
) else (
    echo [警告] 未找到EXE文件，请检查dist目录
)

echo.
echo ============================================================
echo   使用说明
echo ============================================================
echo.
echo 1. 找到生成的 EXE 文件（dist文件夹中）
echo 2. 将 EXE 文件和 Excel 模板放在同一文件夹
echo 3. 双击 EXE 即可运行程序
echo.
echo 首次运行可能需要几秒钟启动，请耐心等待
echo.

pause
