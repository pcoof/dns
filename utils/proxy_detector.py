"""系统代理检测模块"""

import winreg
from typing import Dict, Optional, Tuple
from utils.logger import get_logger

logger = get_logger("ProxyDetector")


class ProxyDetector:
    """系统代理检测器"""

    # 代理相关的注册表路径
    INTERNET_SETTINGS = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"

    @staticmethod
    def check_system_proxy() -> Tuple[bool, Dict[str, Optional[str]]]:
        """
        检查系统是否开启了代理

        Returns:
            tuple: (是否启用代理, 代理信息字典)
                代理信息包含:
                - proxy_enable: 代理启用状态 (0/1)
                - proxy_server: 代理服务器地址
                - proxy_override: 代理绕过列表
                - auto_config_url: 自动配置脚本URL
        """
        try:
            proxy_info = {
                "proxy_enable": None,
                "proxy_server": None,
                "proxy_override": None,
                "auto_config_url": None,
            }

            # 打开注册表
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                ProxyDetector.INTERNET_SETTINGS,
                0,
                winreg.KEY_READ
            ) as key:
                try:
                    # 检查代理启用状态
                    proxy_enable = winreg.QueryValueEx(key, "ProxyEnable")[0]
                    proxy_info["proxy_enable"] = str(proxy_enable)
                except FileNotFoundError:
                    proxy_info["proxy_enable"] = "0"

                try:
                    # 获取代理服务器地址
                    proxy_server = winreg.QueryValueEx(key, "ProxyServer")[0]
                    proxy_info["proxy_server"] = proxy_server
                except FileNotFoundError:
                    pass

                try:
                    # 获取代理绕过列表
                    proxy_override = winreg.QueryValueEx(key, "ProxyOverride")[0]
                    proxy_info["proxy_override"] = proxy_override
                except FileNotFoundError:
                    pass

                try:
                    # 获取自动配置脚本URL
                    auto_config_url = winreg.QueryValueEx(key, "AutoConfigURL")[0]
                    proxy_info["auto_config_url"] = auto_config_url
                except FileNotFoundError:
                    pass

            # 判断是否启用代理
            is_proxy_enabled = (
                proxy_info["proxy_enable"] == "1" or 
                proxy_info.get("auto_config_url") is not None
            )

            if is_proxy_enabled:
                logger.info(f"检测到系统代理已启用: {proxy_info}")
            else:
                logger.info("系统代理未启用")

            return is_proxy_enabled, proxy_info

        except Exception as e:
            logger.error(f"检测系统代理失败: {e}")
            return False, {"error": str(e)}

    @staticmethod
    def get_proxy_warning_message(proxy_info: Dict[str, Optional[str]]) -> str:
        """
        获取代理警告消息

        Args:
            proxy_info: 代理信息字典

        Returns:
            str: 警告消息
        """
        messages = []

        if proxy_info.get("proxy_server"):
            messages.append(f"代理服务器: {proxy_info['proxy_server']}")

        if proxy_info.get("auto_config_url"):
            messages.append(f"自动配置: {proxy_info['auto_config_url']}")

        if proxy_info.get("proxy_override"):
            messages.append(f"绕过列表: {proxy_info['proxy_override']}")

        if messages:
            return "\n".join(messages)
        else:
            return "检测到系统代理已启用，可能影响DNS测试结果"
