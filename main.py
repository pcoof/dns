"""DNS服务器测试工具 - 主入口文件"""

import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb

from config.settings import AppConfig
from ui.main_window import MainWindow
from utils.helpers import is_admin, run_as_admin
from utils.logger import setup_logging, get_logger

# 设置日志
setup_logging(
    log_file=AppConfig.LOG_FILE,
    level=30,  # INFO level
    console_output=True,
)
logger = get_logger("Main")


def show_admin_error():
    """显示管理员权限错误对话框"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "权限错误",
        "DNS测试工具需要管理员权限才能正常运行。\n\n请右键点击程序，选择'以管理员身份运行'。",
        icon="error",
    )
    root.destroy()


def main():
    """主函数"""
    logger.info("=" * 50)
    logger.info("DNS服务器测试工具启动")
    logger.info("=" * 50)

    # 检查管理员权限
    if not is_admin():
        logger.warning("检测到程序未以管理员权限运行")
        logger.info("正在以管理员权限重新启动程序...")

        if run_as_admin():
            logger.info("成功启动管理员权限版本，退出当前进程")
            sys.exit(0)
        else:
            logger.error("无法获取管理员权限")
            show_admin_error()
            sys.exit(1)
    else:
        logger.info("程序已以管理员权限运行")

    try:
        # 创建主窗口（先隐藏，避免闪烁）
        root = tb.Window(themename="darkly")

        # 立即隐藏窗口，在任何显示之前
        root.withdraw()

        # 设置窗口标题
        root.title(f"{AppConfig.WINDOW_TITLE}（管理员权限）")

        # 设置图标
        icon_path = AppConfig.ICON_FILE
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(icon_path)
            except Exception as e:
                logger.warning(f"无法加载图标: {e}")

        # 定义初始化完成回调
        def on_window_ready():
            try:
                current_theme = app.theme_combo.get().lower()
                is_dark = current_theme in AppConfig.DARK_THEMES
                app._set_title_bar_theme(is_dark)
                logger.info(f"初始标题栏主题设置完成: {current_theme}")
            except Exception as e:
                logger.error(f"设置初始标题栏主题失败: {e}")

            # 初始化完成后显示窗口
            root.deiconify()
            logger.info("应用程序初始化完成，窗口已显示")

        # 创建主窗口实例，传入回调
        app = MainWindow(root, on_ready=on_window_ready)

        logger.info("启动主循环")
        root.mainloop()

    except Exception as e:
        logger.exception("应用程序启动失败")
        messagebox.showerror(
            "启动错误",
            f"应用程序启动失败:\n\n{str(e)}\n\n请查看日志文件获取详细信息。",
        )
        sys.exit(1)

    logger.info("应用程序正常退出")


if __name__ == "__main__":
    main()
