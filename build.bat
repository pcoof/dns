@echo off
chcp 65001 >nul
echo 🚀 DNS Tester 构建脚本 (便携版 + 安装版)
echo ===============================================

echo.
echo 📋 构建计划:
echo   1. 构建可执行文件
echo   2. 创建便携版 (包含配置文件)
echo   3. 创建安装程序 (包含配置文件)
echo.

set /p confirm=是否继续? (Y/N): 
if /i not "%confirm%"=="Y" (
    echo 构建已取消
    pause
    exit /b 0
)

echo.
echo ==========================================
echo 🔨 步骤 1: 构建可执行文件
echo ==========================================

call .venv\Scripts\activate
uv sync

echo.
echo 🔨 使用 PyInstaller 构建可执行文件...
uv run pyinstaller --onefile --noconsole --icon=icon.ico --name="DNS-Tester" ^
    --add-data="icon.ico;." ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=ttkbootstrap ^
    --hidden-import=dns ^
    --hidden-import=dns.resolver ^
    --hidden-import=pythoncom ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    main.py

if not exist "dist\DNS-Tester.exe" (
    echo ❌ 可执行文件构建失败！
    pause
    exit /b 1
)

for %%I in ("dist\DNS-Tester.exe") do (
    set /a "exeSize=%%~zI/1024/1024"
    echo ✅ 可执行文件构建成功，大小: !exeSize! MB
)

echo.
echo ==========================================
echo 📦 步骤 2: 创建便携版
echo ==========================================

if exist "DNS-Tester-Portable" rmdir /s /q "DNS-Tester-Portable"
mkdir "DNS-Tester-Portable"

copy "dist\DNS-Tester.exe" "DNS-Tester-Portable\"
copy "dns_servers.ini" "DNS-Tester-Portable\"
copy "README.md" "DNS-Tester-Portable\"
copy "LICENSE" "DNS-Tester-Portable\"

(
echo # DNS Tester 便携版
echo.
echo ## 使用说明
echo 这是 DNS Tester 的便携版本，包含预配置的DNS服务器列表。
echo.
echo ### 文件说明
echo - DNS-Tester.exe - 主程序
echo - dns_servers.ini - DNS服务器配置文件 ^(预配置^)
echo - README.md - 项目说明文档
echo - LICENSE - 许可证文件
echo.
echo ### 使用方法
echo 1. 直接双击 DNS-Tester.exe 运行程序
echo 2. 首次运行可能需要管理员权限
echo 3. 程序会自动加载 dns_servers.ini 中的配置
echo.
echo ### 配置特点
echo - 包含多个分类的DNS服务器
echo - 可以编辑 dns_servers.ini 自定义配置
echo - 所有设置保存在当前目录
echo.
echo 项目地址: https://github.com/pcoof/dns
) > "DNS-Tester-Portable\便携版说明.txt"

powershell -Command "Compress-Archive -Path 'DNS-Tester-Portable\*' -DestinationPath 'DNS-Tester-Portable.zip' -Force"

if exist "DNS-Tester-Portable.zip" (
    for %%I in ("DNS-Tester-Portable.zip") do (
        set /a "zipSize=%%~zI/1024/1024"
        echo ✅ 便携版创建成功，大小: !zipSize! MB
    )
) else (
    echo ❌ 便携版创建失败！
)

echo.
echo ==========================================
echo 🏗️ 步骤 3: 创建安装程序
echo ==========================================

set "innoPath=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "%innoPath%" (
    echo 🔧 使用 Inno Setup 创建安装程序...
    "%innoPath%" installer.iss
    
    if exist "Output\DNS-Tester-Setup.exe" (
        for %%I in ("Output\DNS-Tester-Setup.exe") do (
            set /a "setupSize=%%~zI/1024/1024"
            echo ✅ 安装程序创建成功，大小: !setupSize! MB
        )
    ) else (
        echo ❌ 安装程序创建失败！
    )
) else (
    echo ⚠️ 未找到 Inno Setup，跳过安装程序创建
    echo    请从 https://jrsoftware.org/isinfo.php 下载并安装 Inno Setup
)

echo.
echo ==========================================
echo 🎉 构建完成！
echo ==========================================

echo.
echo 📊 构建结果:

if exist "DNS-Tester-Portable.zip" (
    echo ✅ 便携版: DNS-Tester-Portable.zip ^(!zipSize! MB^)
) else (
    echo ❌ 便携版: 构建失败
)

if exist "Output\DNS-Tester-Setup.exe" (
    echo ✅ 安装版: Output\DNS-Tester-Setup.exe ^(!setupSize! MB^)
) else (
    echo ❌ 安装版: 构建失败或跳过
)

echo.
echo 💡 版本特点:
echo   - 便携版: 解压即用，包含预配置DNS列表
echo   - 安装版: 完整安装体验，包含预配置DNS列表

echo.
echo 🚀 测试程序...
if exist "DNS-Tester-Portable\DNS-Tester.exe" (
    start "" "DNS-Tester-Portable\DNS-Tester.exe"
)

echo.
pause