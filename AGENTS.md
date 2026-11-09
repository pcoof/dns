# DNS Tester — Agent 指南

## 项目概览

Windows DNS 服务器测试工具。GUI 基于 `tkinter` + `ttkbootstrap`，网络操作使用 WMI COM（`pywin32`），DNS 查询使用 `dnspython`。仅支持 Windows 10/11。

## 架构

```
main.py               → 入口（权限检查 → 创建窗口 → 主循环）
config/               → settings.py（常量）、paths.py（运行时路径解析）、
                         models.py（dataclass）、config_manager.py（INI 读写）
core/                 → dns_tester.py（DNSTester，并行测试）、
                         network_manager.py（WMI 网络操作）、
                         dns_config_handler.py（配置业务逻辑）
ui/                   → main_window.py（界面）、dialogs.py（对话框）
utils/                → helpers.py、logger.py、wmi_context.py、
                         history_manager.py、export_manager.py、proxy_detector.py
assets/               → icon.ico（程序图标）、portable_readme.txt（便携版说明）
packaging/            → installer.iss（Inno Setup 安装脚本，UTF-8 with BOM）
docs/                 → BUILD_VERSIONS.md、OPTIMIZATION_SUMMARY.md、RELEASE_TEMPLATE.md
tests/                → pytest 单元测试（mock 模式，无集成测试）
.github/workflows/    → build.yml（质量检查 → 构建 → Release）
```

**关键事实**：项目已从单文件 `main.py`（原 2127 行）重构为模块化结构。

## 开发命令

```bash
# 依赖管理（必须使用 uv，项目锁定 uv.lock）
uv sync                     # 安装运行依赖
uv sync --extra build       # 追加 PyInstaller
uv sync --extra dev         # 追加 pytest/ruff/black/mypy
uv run python main.py       # 运行（需管理员终端）

# 代码质量（配置在 pyproject.toml 中）
ruff check .                # 静态检查，line-length=100
black .                     # 格式化，line-length=100
mypy .                      # 类型检查，strict 模式，ignore_missing_imports=true

# 测试
uv run pytest -q            # 运行所有测试
uv run pytest -q --no-cov   # 跳过覆盖率插件（CI 使用）
```

本地手工打包（CI 会做同样的事，一般不必本地执行）：

```bash
uv run pyinstaller --noconfirm --clean --onefile --noconsole --icon "assets\icon.ico" --add-data "assets\icon.ico;assets" --name "DNS-Tester" main.py
```

## 平台与权限

- **仅 Windows**：WMI COM、注册表操作、管理员提权均通过 `ctypes`/`pywin32` 实现。
- **强制管理员权限**：`main.py` 启动时调用 `is_admin()` 检查，非管理员则自动通过 `ShellExecuteW("runas")` 提权重启。**调试/运行时必须以管理员身份运行**，否则程序会闪退。
- **WMI 上下文**：`utils/wmi_context.py` 提供两个上下文管理器：
  - `wmi_context()` — 使用 `winmgmts:\\\\.\\` 连接（只读操作用）
  - `wmi_dispatch_context()` — 使用 `WbemScripting.SWbemLocator` 连接（写操作如 `SetDNSServerSearchOrder` 用）
  - 两者都自动处理 `pythoncom.CoInitialize()` / `CoUninitialize()`。

## 配置与路径

- **路径解析**：`config/paths.py` 是唯一的路径来源，不要在业务代码里写相对路径或 `os.path.dirname(__file__)`。
  - `APP_DIR` — 可写数据目录。打包后为 exe 所在目录，源码运行为项目根。
  - `RESOURCE_DIR` — 只读资源目录。onefile 模式下为 `sys._MEIPASS`。
  - `resource_path("assets/icon.ico")` / `data_path("dns_config.ini")`
  - 该模块**不得**导入项目内其他模块，否则会与 `config.settings` 形成循环导入。
- **格式**：INI（`configparser`），文件名 `dns_config.ini`，位于 `APP_DIR`。
- **分类**：节名格式 `Ipv4_*` / `Ipv6_*`。保留节 `[Main]` 存储元数据（`theme`、`last_category`、`category_order`）。
- **DNS 条目格式**：`name = primary_dns,secondary_dns[,True]`（第三个值当前代码读取但未使用）。
- **无配置文件时**：`ConfigManager.create_default_config()` 自动创建精简默认列表；发行版随包提供完整预置列表。
- **ConfigManager** 加载时设置了 `optionxform = str`（区分大小写）。
- **运行时产物不入库**：`dns_config.ini` 作为预置模板保留在仓库，`dns_history.json`、`dns_tester.log` 由 `.gitignore` 排除。

## 已知问题（Agent 请注意）

### 功能差距
- **`utils/export_manager.py`（ExportManager）已实现但未集成到 UI**。`export_dns_to_csv`、`export_dns_to_json`、`import_dns_from_json`、`import_dns_from_csv` 均无调用方。如需接入，连接到 UI 的按钮/菜单项。

### 历史遗留
- `docs/OPTIMIZATION_SUMMARY.md`、`docs/BUILD_VERSIONS.md` 是重构期记录，部分描述的是旧流程，仅作历史参考，不作为当前架构依据。
- `.kiro/` 目录已删除。

## 编码约定

- 注释使用中文，变量/函数名使用英文。
- 全局字体：`Microsoft YaHei`。
- 编码：所有文件读写使用 `utf-8`。INI 文件也强制 `utf-8`。`packaging/installer.iss` 必须带 UTF-8 BOM，否则 Inno Setup 会按 ANSI 解析导致中文乱码。
- 类型注解：项目广泛使用类型注解，mypy strict 模式已开启，类型错误**不得**用 `# type: ignore` 绕过。
- `ruff check .` 必须零告警（CI 的 quality job 会强制校验）。

## 构建与发布

- **工作流**：`.github/workflows/build.yml`
  - `quality`：`ruff check` + `pytest`（PR、push、tag 均触发）
  - `build`：`uv sync --frozen --extra build` → PyInstaller → 便携版 zip → Inno Setup 安装包 → SHA256 校验和 → artifact
  - Release 仅在推送 `v*` 标签（或手动勾选 `create_release`）时创建
- **Inno Setup**：runner 预装于 `C:\Program Files (x86)\Inno Setup 6\ISCC.exe`，缺失时用 Chocolatey 兜底。
- **版本号**：tag 构建取 tag 名（去掉 `v`），否则取 `pyproject.toml` 的 `version`。
- **发布流程**：
  ```bash
  git tag v2.0.1
  git push origin v2.0.1
  ```
- **CI 排版注意**：`run: |` 块内的 PowerShell here-string（`@" ... "@`）正文若从第 0 列开始，会提前终止 YAML 块标量，导致整份工作流解析失败（表现为 run 耗时 0 秒、没有任何 job）。需要多行文本时改用 `@(...) -join "`n"` 或直接引用仓库内的文件。

## 测试注意事项

- 所有测试基于 mock（`unittest.mock`），不依赖真实网络或 WMI。
- `test_config_manager.py` 使用 `tempfile.NamedTemporaryFile` 创建临时 INI 文件。
- `test_dns_tester.py` mock 了 `dns.resolver.Resolver`。
- 无集成测试。对 WMI 或真实 DNS 解析器的改动**必须手动在 Windows 上以管理员身份测试**。

## 代理检测

- `utils/proxy_detector.py` 通过 Windows 注册表（`HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings`）检查系统代理状态。
- 测试开始时，若系统代理开启则显示警告对话框，这是设计行为，**不是 Bug**。

## 关键依赖

| 包 | 用途 | 版本 |
|---|---|---|
| `dnspython` | DNS 查询 | ==2.7.0 |
| `ttkbootstrap` | GUI 主题 | ==1.14.0 |
| `pywin32` | WMI COM 接口 | ==311 |
| `pyinstaller` | 打包（`build` extra） | >=6.14.2 |

`pyproject.toml` 中 `dev` extra 包含 `pytest`、`pytest-cov`、`black`、`mypy`、`ruff` 的版本约束。
