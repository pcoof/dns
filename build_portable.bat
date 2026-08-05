@echo off
chcp 65001 >nul
echo 📦 构建 DNS Tester 便携版 (包含配置文件)

if not exist "dist\DNS-Tester.exe" (
    echo ❌ 找不到可执行文件！
    echo 请先运行 build_single.bat 构建单文件版
    pause
    exit /b 1
)

echo.
echo 📁 创建便携版目录...
if exist "DNS-Tester-Portable" rmdir /s /q "DNS-Tester-Portable"
mkdir "DNS-Tester-Portable"

echo.
echo 📋 复制文件...
copy "dist\DNS-Tester.exe" "DNS-Tester-Portable\"
copy "dns_servers.ini" "DNS-Tester-Portable\"
copy "README.md" "DNS-Tester-Portable\"
copy "LICENSE" "DNS-Tester-Portable\"

echo.
echo 📝 创建便携版说明文件...
(
echo # DNS Tester 便携版
echo.
echo ## 使用说明
echo.
echo 这是 DNS Tester 的便携版本，包含预配置的DNS服务器列表。
echo.
echo ### 文件说明
echo - DNS-Tester.exe - 主程序
echo - dns_servers.ini - DNS服务器配置文件 ^(预配置^)
echo - README.md - 项目说明文档
echo - LICENSE - 许可证文件
echo - 便携版说明.txt - 本文件
echo.
echo ### 使用方法
echo 1. 直接双击 DNS-Tester.exe 运行程序
echo 2. 首次运行可能需要管理员权限（用于修改系统DNS设置）
echo 3. 程序会自动加载 dns_servers.ini 中的DNS服务器配置
echo 4. 所有设置和配置会保存在当前目录下
echo.
echo ### 配置文件
echo - dns_servers.ini 包含了多个分类的DNS服务器
echo - 你可以编辑此文件来添加或修改DNS服务器
echo - 程序启动时会自动加载这些配置
echo.
echo ### 系统要求
echo - Windows 10/11 ^(64位^)
echo - 无需安装任何运行时或依赖
echo.
echo ### 注意事项
echo - 便携版可以放在U盘等移动存储设备中使用
echo - 建议将整个文件夹复制到本地硬盘使用以获得最佳性能
echo - 如需卸载，直接删除整个文件夹即可
echo.
echo ---
echo 项目地址: https://github.com/pcoof/dns
) > "DNS-Tester-Portable\便携版说明.txt"

echo.
echo 🗜️ 创建ZIP压缩包...
powershell -Command "Compress-Archive -Path 'DNS-Tester-Portable\*' -DestinationPath 'DNS-Tester-Portable.zip' -Force"

if exist "DNS-Tester-Portable.zip" (
    echo ✅ 便携版构建完成！
    echo.
    echo 📊 文件信息:
    for %%I in ("DNS-Tester-Portable.zip") do (
        set /a "size=%%~zI/1024/1024"
        echo    压缩包大小: !size! MB
    )
    echo    📁 便携版目录: DNS-Tester-Portable\
    echo    📦 压缩包文件: DNS-Tester-Portable.zip
    echo    📝 特点: 包含预配置的DNS服务器列表
    
    echo.
    echo 🚀 打开便携版目录...
    start "" "DNS-Tester-Portable"
) else (
    echo ❌ 便携版构建失败！
)

echo.
pause