# 项目结构和组织

## 文件结构

```
dns-tester/
├── .git/                    # Git版本控制
├── .kiro/                   # Kiro AI助手配置
│   └── steering/            # AI助手指导规则
├── .venv/                   # Python虚拟环境
├── dist/                    # PyInstaller打包输出目录
├── main.py                  # 主应用程序文件 (单文件架构)
├── dns_servers.ini          # DNS服务器配置文件
├── icon.ico                 # 应用程序图标
├── pyproject.toml          # 项目配置和依赖
├── uv.lock                 # 依赖锁定文件
├── .python-version         # Python版本指定
├── .gitignore              # Git忽略规则
└── README.md               # 项目文档
```

## 核心文件说明

### main.py
- **单文件应用程序**: 包含完整的GUI应用逻辑
- **主要类**: `DNSTesterApp` - 主应用程序类
- **功能模块**:
  - GUI界面创建和管理
  - DNS测试逻辑
  - 网络适配器检测
  - 配置文件管理
  - 系统DNS设置

### dns_servers.ini
- **配置格式**: INI文件格式
- **分类结构**: 
  - `[Main]` - 应用程序设置
  - `[Ipv4_*]` - IPv4 DNS服务器分类
  - `[Ipv6_*]` - IPv6 DNS服务器分类
- **DNS条目格式**: `name = primary_dns,secondary_dns`

## 代码组织原则

### 类结构
- **单一职责**: `DNSTesterApp`类负责整个应用程序
- **方法分组**:
  - UI创建: `create_widgets()`, `show_*()` 方法
  - 网络操作: `get_*()`, `load_network_*()` 方法  
  - DNS测试: `test_*()`, `start_test()` 方法
  - 配置管理: `load_*()`, `save_*()` 方法

### 命名约定
- **中文注释**: 所有注释使用中文
- **英文变量名**: 变量和方法名使用英文
- **描述性命名**: 方法名清楚表达功能用途

### 文件管理
- **配置持久化**: 自动保存用户设置到INI文件
- **错误处理**: 优雅处理文件读写错误
- **编码支持**: 统一使用UTF-8编码处理中文

## 构建输出

### dist/ 目录
- **可执行文件**: PyInstaller生成的独立exe文件
- **依赖打包**: 所有Python依赖库打包在内
- **资源文件**: 图标等资源文件嵌入