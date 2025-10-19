@echo off
setlocal enabledelayedexpansion

echo DNS Tester 发布脚本
echo.

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
git status --porcelain > nul
if errorlevel 1 (
    echo 错误: Git仓库有未提交的更改
    pause
    exit /b 1
)

echo 2. 创建并推送标签...
git tag %VERSION%
git push origin %VERSION%

echo.
echo 3. 标签已创建并推送到GitHub
echo GitHub Actions 将自动开始构建...
echo.
echo 请访问以下链接查看构建进度:
echo https://github.com/pcoof/dns/actions
echo.

pause