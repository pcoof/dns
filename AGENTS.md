# DNS Tester — Agent 指南

## 项目概览

Windows DNS 服务器测试工具。GUI 基于 `tkinter` + `ttkbootstrap`，网络操作使用 WMI COM（`pywin32`），DNS 查询使用 `dnspython`。仅支持 Windows 10/11。

## 架构

```
main.py               → 入口（~108 行，权限检查 → 创建窗口 → 主循环）
config/               → settings.py（常量）、models.py（dataclass）、config_manager.py（INI 读写）
core/                 → dns_tester.py（DNSTester，并行测试）、network_manager.py（WMI 网络操作）、dns_config_handler.py（配置业务逻辑）
ui/                   → main_window.py（界面）、dialogs.py（对话框）
utils/                → helpers.py、logger.py、wmi_context.py、history_manager.py、export_manager.py、proxy_detector.py
tests/                → pytest 单元测试（mock 模式，无集成测试）
```

**关键事实**：项目已从单文件 `main.py`（原 2127 行）重构为模块化结构。旧版本备份在 `main_backup.py`。

## 开发命令

```bash
# 依赖管理（必须使用 uv，项目锁定 uv.lock）
uv sync                     # 安装依赖
uv sync --frozen            # CI 中使用，不更新锁文件
uv run python main.py       # 运行

# 代码质量（配置在 pyproject.toml 中）
black .                     # 格式化，line-length=100
mypy .                      # 类型检查，strict 模式，ignore_missing_imports=true
ruff .                      # 静态检查，line-length=100

# 测试
pytest tests/ -v            # 运行所有测试
pytest tests/test_helpers.py -v   # 指定模块
pytest --cov=. --cov-report=html  # 覆盖率报告

# 构建（Windows 专用）
build.bat                   # 完整构建（exe → 便携版 zip → Inno Setup 安装包）
build_portable.bat          # 仅构建便携版
```

## 平台与权限

- **仅 Windows**：WMI COM、注册表操作、管理员提权均通过 `ctypes`/`pywin32` 实现。
- **强制管理员权限**：`main.py` 启动时调用 `is_admin()` 检查，非管理员则自动通过 `ShellExecuteW("runas")` 提权重启。**调试/运行时必须以管理员身份运行**，否则程序会闪退。
- **WMI 上下文**：`utils/wmi_context.py` 提供两个上下文管理器：
  - `wmi_context()` — 使用 `winmgmts:\\\\.\\` 连接（只读操作用）
  - `wmi_dispatch_context()` — 使用 `WbemScripting.SWbemLocator` 连接（写操作如 `SetDNSServerSearchOrder` 用）
  - 两者都自动处理 `pythoncom.CoInitialize()` / `CoUninitialize()`。

## 配置系统

- **格式**：INI（`configparser`），文件名 `dns_servers.ini`。
- **分类**：节名格式 `Ipv4_*` / `Ipv6_*`。保留节 `[Main]` 存储元数据（`theme`、`last_category`、`category_order`）。
- **DNS 条目格式**：`name = primary_dns,secondary_dns[,True]`（第三个值 `True` 当前代码读取但未使用，仅作保留）。
- **配置位置**：程序运行目录。无配置文件时，`ConfigManager.create_default_config()` 会自动创建。
- **ConfigManager** 加载时设置了 `optionxform = str`（区分大小写）。

## 已知问题（Agent 请注意）

### Bug
- **`ui/main_window.py` 中的 `_start_dns_test` 方法被复制了两次**（第 369 行和第 390 行）。第二个覆盖了第一个，所以第一个实际上不可达。修改时请使用第二个（从第 390 行开始的版本）。

### 功能差距
- **`utils/export_manager.py`（ExportManager）已实现但未集成到 UI 中**。其 `export_dns_to_csv`、`export_dns_to_json`、`import_dns_from_json`、`import_dns_from_csv` 方法均未在任何地方被调用。如需接入，连接到 UI 中的按钮/菜单项。

### .kiro/ 目录
- `.kiro/steering/` 下的文件（`product.md`、`structure.md`、`tech.md`）由另一个 AI 工具（kiroAgent）生成，**描述的是旧版单文件架构**，已过时。不可将其视为当前架构的权威参考。

## 编码约定

- 注释使用中文，变量/函数名使用英文。
- 全局字体：`Microsoft YaHei`。
- 编码：所有文件读写使用 `utf-8`。INI 文件也强制 `utf-8`。
- 类型注解：项目广泛使用类型注解，mypy strict 模式已开启，类型错误**不得**用 `# type: ignore` 绕过。

## 构建与发布

- **构建工具**：PyInstaller（`--onefile --noconsole`），许多 `--hidden-import` 是必需的（`tkinter`、`ttkbootstrap`、`dns.resolver`、`pythoncom`、`win32api`、`win32con`、`dns`）。
- **.spec 文件**：`DNS-Tester.spec` 与 `pyproject.toml` 中的 `pyinstaller` 构建命令并存。CI 使用内联 PyInstaller 命令而非 .spec 文件。
- **发布流程**：
  - `release.bat v1.0.0` → 创建 git tag 并推送 → GitHub Actions（`build.yml`）自动构建并创建 Release。
  - CI 配置在 `.github/workflows/build.yml` 中。它会构建 exe、便携版 zip 和 Inno Setup 安装包，计算 SHA256 校验和，并将产物上传到 Release。
  - Inno Setup 安装在 CI runner 的 `C:\Program Files (x86)\Inno Setup 6\ISCC.exe` 下。

## 测试注意事项

- 所有测试基于 mock（`unittest.mock`），不依赖真实网络或 WMI。
- `test_config_manager.py` 使用 `tempfile.NamedTemporaryFile` 创建临时 INI 文件。
- `test_dns_tester.py` mock 了 `dns.resolver.Resolver`。
- 无集成测试。对 WMI 或真实 DNS 解析器的改动**必须手动在 Windows 上以管理员身份测试**。
- 测试覆盖率通过 `pytest-cov` 生成（`--cov=. --cov-report=html`）。

## 代理检测

- `utils/proxy_detector.py` 通过 Windows 注册表（`HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings`）检查系统代理状态。
- 测试开始时，若系统代理开启则显示警告对话框，这是设计行为，**不是 Bug**。

## 关键依赖

| 包 | 用途 | 版本 |
|---|---|---|
| `dnspython` | DNS 查询 | ==2.7.0 |
| `ttkbootstrap` | GUI 主题 | ==1.14.0 |
| `pywin32` | WMI COM 接口 | ==311 |
| `pyinstaller` | 打包（可选依赖） | >=6.14.2 |

`pyproject.toml` 中 dev 依赖包含 `pytest`、`pytest-cov`、`black`、`mypy`、`ruff` 的版本约束。
