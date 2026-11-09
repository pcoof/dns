"""日志系统配置模块"""

import logging
from pathlib import Path


def setup_logging(
    log_file: str = "dns_tester.log",
    level: int = logging.INFO,
    console_output: bool = True,
) -> logging.Logger:
    """
    配置应用程序日志系统

    Args:
        log_file: 日志文件路径
        level: 日志级别 (logging.DEBUG, logging.INFO, logging.WARNING, etc.)
        console_output: 是否输出到控制台

    Returns:
        配置好的logger实例

    Example:
        >>> logger = setup_logging(log_file="app.log", level=logging.INFO)
        >>> logger.info("应用程序启动")
    """
    # 创建logger
    logger = logging.getLogger("DNS Tester")
    logger.setLevel(level)

    # 清除现有的处理器
    logger.handlers.clear()

    # 日志格式
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 文件处理器
    try:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        print(f"无法创建日志文件处理器: {e}")

    # 控制台处理器
    if console_output:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


def get_logger(name: str) -> logging.Logger:
    """
    获取指定名称的logger

    Args:
        name: logger名称

    Returns:
        logger实例
    """
    return logging.getLogger(f"DNS Tester.{name}")