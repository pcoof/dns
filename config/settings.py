"""应用程序配置常量模块"""

from typing import List, Dict


class AppConfig:
    """应用程序配置类"""

    # 窗口配置
    WINDOW_TITLE: str = "DNS服务器测试工具"
    WINDOW_SIZE: tuple = (980, 500)
    FONT_FAMILY: str = "Microsoft YaHei"

    # 文件配置
    CONFIG_FILE: str = "dns_config.ini"
    HISTORY_FILE: str = "dns_history.json"
    LOG_FILE: str = "dns_tester.log"

    # 主题配置
    THEMES: List[str] = [
        "Cyborg",
        "Darkly",
        "Superhero",
        "Vapor",
        "Solar",
        "Flatly",
        "Cosmo",
        "Journal",
        "Litera",
        "Lumen",
        "Minty",
        "Morph",
        "Pulse",
        "Sandstone",
        "Simplex",
        "Yeti",
        "United",
    ]

    # 深色主题列表
    DARK_THEMES: set = {
        "cyborg",
        "darkly",
        "superhero",
        "vapor",
        "solar",
    }

    # 默认DNS服务器配置
    DEFAULT_IPV4: List[Dict[str, str]] = [
        {"name": "CN - 阿里云", "primary": "223.5.5.5", "secondary": "223.6.6.6"},
        {"name": "CN - 114DNS","primary": "114.114.114.114","secondary": "114.114.115.115"},
        {"name": "CN - DNSPod","primary": "119.29.29.29","secondary": "182.254.116.116"},
    ]

    DEFAULT_IPV6: Dict[str, str] = {
        "Google": "2001:4860:4860::8888,2001:4860:4860::8844",
        "阿里云": " 2400:3200::1,2400:3200:baba::1", 
        "百度云": "2400:da00::6666",
    }

    # DNS测试配置
    DNS_TEST_DOMAIN: str = "www.baidu.com"
    DNS_TEST_TIMEOUT: int = 3  # 秒
    DNS_TEST_MAX_WORKERS: int = 5  # 并发测试线程数

    # UI配置
    TREEVIEW_ROW_HEIGHT: int = 45
    TREEVIEW_COLUMNS: List[str] = [
        "name",
        "primary",
        "secondary",
        "latency",
        "status",
    ]

    # 保留的配置节名称
    RESERVED_SECTION_NAME: str = "Main"

# https://dns.alidns.com/dns-query
# https://dns.adguard-dns.com/dns-query
# https://cloudflare-dns.com/dns-query
# https://dnsdaquan.com/


