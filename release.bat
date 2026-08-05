@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo 🚀 DNS Tester 发布脚本
echo =======================

if "%1"=="" (
    echo 用法: release.bat [版本号]
    echo 例如: release.bat v1.0.0
    echo.
    pause
    exit /b 1
)

set VERSION=%1

echo 准备发布版本: %VERSION%
echo.

echo 1. 检查Git状态...
git status --porcelain > nul 2>&1
if errorlevel 1 (
    echo ❌ Git仓库有未提交的更改
    pause
    exit /b 1
)

echo 2. 创建并推送标签...
git tag %VERSION%
if errorlevel 1 (
    echo ❌ 创建标签失败
    pause
    exit /b 1
)

git push origin %VERSION%
if errorlevel 1 (
    echo ❌ 推送标签失败
    pause
    exit /b 1
)

echo.
echo ✅ 标签 %VERSION% 已创建并推送到GitHub
echo.
echo 📦 GitHub Actions 将自动构建以下版本:
echo   - 便携版 (DNS-Tester-Portable.zip)
echo   - 安装版 (DNS-Tester-Setup.exe)
echo.
echo 🔗 请访问以下链接查看构建进度:
echo    https://github.com/pcoof/dns/actions
echo.
echo 📋 构建完成后，Release将包含:
echo   - DNS-Tester-Portable.zip - 便携版，包含预配置DNS列表
echo   - DNS-Tester-Setup.exe - 安装版，包含预配置DNS列表
echo   - checksums.txt - 文件校验和
echo.

pause