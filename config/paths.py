"""运行时路径解析模块

统一处理三种运行场景下的路径定位：

1. 源码运行（``uv run python main.py``）：以项目根目录为基准
2. PyInstaller onefile：可写数据放在 exe 所在目录，只读资源在 ``sys._MEIPASS``
3. PyInstaller onedir：同上，可写数据在 exe 目录

本模块不允许导入任何项目内其他模块，避免与 ``config.settings`` 形成循环导入。
"""

import sys
from pathlib import Path


def _frozen() -> bool:
    """是否为 PyInstaller 打包后的运行环境"""
    return bool(getattr(sys, "frozen", False))


def get_app_dir() -> Path:
    """返回可写数据目录（配置文件、历史记录、日志所在目录）

    - 打包后：exe 所在目录（便携版解压目录 / 安装目录）
    - 源码运行：项目根目录
    """
    if _frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_resource_dir() -> Path:
    """返回只读资源目录（随程序分发的静态文件）

    onefile 模式下 PyInstaller 会把 ``--add-data`` 的内容解压到临时目录，
    该路径通过 ``sys._MEIPASS`` 暴露。
    """
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent


APP_DIR: Path = get_app_dir()
RESOURCE_DIR: Path = get_resource_dir()


def resource_path(relative: str) -> Path:
    """将相对资源路径转换为绝对路径

    Args:
        relative: 相对于资源根目录的路径，如 ``"assets/icon.ico"``

    Returns:
        绝对路径
    """
    return RESOURCE_DIR / relative


def data_path(relative: str) -> Path:
    """将相对数据路径转换为绝对路径

    Args:
        relative: 相对于可写数据目录的路径，如 ``"dns_config.ini"``

    Returns:
        绝对路径
    """
    return APP_DIR / relative
