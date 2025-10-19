@echo off
echo 正在构建 DNS Tester...

echo.
echo 1. 激活虚拟环境...
call .venv\Scripts\activate

echo.
echo 2. 安装/更新依赖...
uv sync

echo.
echo 3. 使用 PyInstaller 构建可执行文件...
uv run pyinstaller --onefile --windowed --icon=icon.ico --name="DNS-Tester" main.py

echo.
echo 4. 构建完成！
echo 可执行文件位置: dist\DNS-Tester.exe

echo.
echo 5. 如需创建安装包，请确保已安装 Inno Setup，然后运行:
echo "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss

pause