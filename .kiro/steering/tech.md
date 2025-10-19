# 技术栈和构建系统

## 技术栈

- **Python**: 3.13+ (主要编程语言)
- **GUI框架**: tkinter + ttkbootstrap (现代化UI主题)
- **DNS库**: dnspython (DNS查询和解析)
- **打包工具**: PyInstaller (生成独立可执行文件)
- **配置管理**: configparser (INI文件处理)
- **依赖管理**: uv (Python包管理器)

## 项目依赖

```toml
dependencies = [
    "dnspython>=2.7.0",
    "pyinstaller>=6.14.2", 
    "ttkbootstrap>=1.14.0",
]
```

## 常用命令

### 开发环境设置
```bash
# 激活虚拟环境 (Windows)
.venv\Scripts\activate

# 安装依赖
uv sync
```

### 运行应用
```bash
python main.py
```

### 打包应用
```bash
# 使用PyInstaller打包为单个可执行文件
pyinstaller --onefile --windowed --icon=icon.ico main.py

# 打包后的文件位于 dist/ 目录
```

## 平台特定

- **目标平台**: Windows 10/11
- **权限要求**: 设置系统DNS时需要管理员权限
- **编码**: 支持中文界面，使用UTF-8编码
- **系统集成**: 通过Windows命令行工具(ipconfig, route)获取网络信息

## 开发约定

- 单文件应用程序架构
- 中文注释和文档
- 现代化UI设计原则
- 错误处理和用户友好的提示信息