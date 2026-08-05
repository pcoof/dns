"""网络管理器模块 - 处理网络适配器和DNS配置"""

import platform
import socket
import subprocess
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

from utils.wmi_context import wmi_context, wmi_dispatch_context
from utils.helpers import is_ip
from utils.logger import get_logger

logger = get_logger("NetworkManager")


@dataclass
class NetworkAdapter:
    """网络适配器数据类"""

    name: str  # 友好名称（如"以太网"）
    index: int  # 适配器索引
    mac: str  # MAC地址
    ipv4_addresses: List[str]  # IPv4地址列表
    ipv6_addresses: List[str]  # IPv6地址列表
    ipv4_gateways: List[str]  # IPv4网关列表
    ipv6_gateways: List[str]  # IPv6网关列表
    dns_servers: List[str]  # DNS服务器列表
    dhcp_enabled: bool  # 是否启用DHCP


class NetworkManager:
    """网络适配器和DNS配置管理器"""

    def __init__(self):
        """初始化网络管理器"""
        self.platform = platform.system()
        logger.info(f"网络管理器初始化，操作系统: {self.platform}")

    def get_network_adapters(self) -> List[NetworkAdapter]:
        """
        获取所有启用的网络适配器信息

        Returns:
            网络适配器列表

        Raises:
            Exception: 当获取网络适配器信息失败时
        """
        if self.platform != "Windows":
            logger.warning("此功能仅支持Windows系统")
            return []

        try:
            with wmi_context() as wmi:
                # 获取启用的网络适配器
                adapters = wmi.ExecQuery(
                    "SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True"
                )

                # 创建适配器信息映射
                adapter_info_map = {}
                for adapter in adapters:
                    adapter_info_map[str(adapter.DeviceID)] = {
                        "name": adapter.Name,
                        "net_connection_id": adapter.NetConnectionID,
                    }

                # 获取网络适配器配置
                configs = wmi.ExecQuery(
                    "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
                )

                network_adapters = []
                for config in configs:
                    adapter_info = adapter_info_map.get(str(config.Index), {})

                    # 优先使用NetConnectionID（友好名称）
                    adapter_name = (
                        adapter_info.get("net_connection_id")
                        or adapter_info.get("name")
                        or "未识别适配器"
                    )

                    # 分离IPv4和IPv6地址
                    ip_addresses = list(config.IPAddress) if config.IPAddress else []
                    ipv4_addresses = [
                        ip for ip in ip_addresses if not is_ip(ip, 6)
                    ]
                    ipv6_addresses = [ip for ip in ip_addresses if is_ip(ip, 6)]

                    # 分离IPv4和IPv6网关
                    gateways = (
                        list(config.DefaultIPGateway)
                        if config.DefaultIPGateway
                        else []
                    )
                    ipv4_gateways = [gw for gw in gateways if not is_ip(gw, 6)]
                    ipv6_gateways = [gw for gw in gateways if is_ip(gw, 6)]

                    adapter = NetworkAdapter(
                        name=adapter_name,
                        index=config.Index,
                        mac=config.MACAddress,
                        ipv4_addresses=ipv4_addresses,
                        ipv6_addresses=ipv6_addresses,
                        ipv4_gateways=ipv4_gateways,
                        ipv6_gateways=ipv6_gateways,
                        dns_servers=list(config.DNSServerSearchOrder)
                        if config.DNSServerSearchOrder
                        else [],
                        dhcp_enabled=config.DHCPEnabled,
                    )
                    network_adapters.append(adapter)

                logger.info(f"获取到 {len(network_adapters)} 个网络适配器")
                return network_adapters

        except Exception as e:
            logger.error(f"获取网络适配器信息失败: {e}")
            raise

    def get_network_connections(self) -> List[str]:
        """
        获取可用的网络连接名称列表

        Returns:
            网络连接名称列表
        """
        try:
            adapters = self.get_network_adapters()
            connections = [
                adapter.name
                for adapter in adapters
                if adapter.ipv4_addresses or adapter.ipv6_addresses
            ]
            logger.info(f"获取到 {len(connections)} 个网络连接")
            return connections if connections else ["未找到网络连接"]
        except Exception as e:
            logger.error(f"获取网络连接失败: {e}")
            # 备选方案
            return ["以太网", "WLAN", "Wi-Fi"]

    def get_local_ip(self) -> str:
        """
        获取本机局域网IP地址

        Returns:
            本机IP地址，如果获取失败返回"无法获取"
        """
        try:
            # 使用WMI获取网络适配器信息
            adapters = self.get_network_adapters()
            for adapter in adapters:
                if adapter.ipv4_addresses:
                    for ip in adapter.ipv4_addresses:
                        # 排除回环地址和APIPA地址
                        if not ip.startswith(("127.", "169.254.")):
                            logger.info(f"获取到本机IP: {ip}")
                            return ip

            # 备选方案：使用socket方法
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                s.connect(("8.8.8.8", 80))
                local_ip = s.getsockname()[0]
                logger.info(f"通过socket获取到本机IP: {local_ip}")
                return local_ip

        except Exception as e:
            logger.error(f"获取本机IP失败: {e}")
            try:
                # 最后备选方案
                hostname = socket.gethostname()
                local_ip = socket.gethostbyname(hostname)
                logger.info(f"通过主机名获取到本机IP: {local_ip}")
                return local_ip
            except Exception:
                logger.error("所有方法均无法获取本机IP")
                return "无法获取"

    def get_default_gateway(self) -> str:
        """
        获取默认网关地址

        Returns:
            默认网关地址，如果获取失败返回"无法获取"
        """
        try:
            adapters = self.get_network_adapters()
            for adapter in adapters:
                if adapter.ipv4_gateways:
                    gateway = adapter.ipv4_gateways[0]
                    logger.info(f"获取到默认网关: {gateway}")
                    return gateway
        except Exception as e:
            logger.error(f"获取默认网关失败: {e}")
        return "无法获取"

    def get_current_dns_servers(self) -> List[str]:
        """
        获取当前系统DNS服务器

        Returns:
            DNS服务器列表，确保返回2个元素
        """
        try:
            adapters = self.get_network_adapters()
            for adapter in adapters:
                if adapter.dns_servers:
                    dns_list = adapter.dns_servers[:2]
                    # 确保返回两个元素
                    if len(dns_list) == 1:
                        dns_list.append("未设置")
                    logger.info(f"获取到DNS服务器: {dns_list}")
                    return dns_list
        except Exception as e:
            logger.error(f"获取当前DNS服务器失败: {e}")
        return ["无法获取", ""]

    def set_dns_servers(
        self,
        adapter_name: str,
        dns_servers: List[str],
        enable_dhcp: bool = False,
    ) -> bool:
        """
        设置指定网络适配器的DNS服务器

        Args:
            adapter_name: 网络适配器名称
            dns_servers: DNS服务器列表
            enable_dhcp: 是否启用DHCP自动获取DNS

        Returns:
            bool: 设置成功返回True，失败返回False

        Raises:
            Exception: 当设置DNS失败时
        """
        if self.platform != "Windows":
            logger.error("此功能仅支持Windows系统")
            return False

        try:
            logger.info(
                f"开始设置DNS - 适配器: {adapter_name}, "
                f"DNS: {dns_servers}, DHCP: {enable_dhcp}"
            )

            with wmi_dispatch_context() as wmi:
                # 获取所有启用的网络适配器配置
                configs = wmi.ExecQuery(
                    "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
                )

                # 获取网络适配器信息以匹配名称
                adapters = wmi.ExecQuery(
                    "SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True"
                )
                adapter_map = {}
                for adapter in adapters:
                    adapter_map[str(adapter.DeviceID)] = {
                        "name": adapter.Name,
                        "net_connection_id": adapter.NetConnectionID,
                    }

                # 查找匹配的适配器配置
                target_config = None
                for config in configs:
                    adapter_info = adapter_map.get(str(config.Index), {})
                    friendly_name = adapter_info.get("net_connection_id", "")
                    adapter_full_name = adapter_info.get("name", "")

                    if (
                        friendly_name == adapter_name
                        or adapter_full_name == adapter_name
                        or adapter_name in friendly_name
                        or adapter_name in adapter_full_name
                    ):
                        target_config = config
                        logger.info(
                            f"找到匹配的网络适配器: {friendly_name or adapter_full_name}"
                        )
                        break

                if not target_config:
                    logger.error(f"未找到网络适配器: {adapter_name}")
                    return False

                # 调用WMI方法设置DNS
                try:
                    if enable_dhcp:
                        logger.info("设置DNS为DHCP自动获取")
                        method_params = (
                            target_config.Methods_("SetDNSServerSearchOrder")
                            .InParameters.SpawnInstance_()
                        )
                        method_params.DNSServerSearchOrder = []
                        result = target_config.ExecMethod_(
                            "SetDNSServerSearchOrder", method_params
                        )
                    else:
                        if not dns_servers:
                            logger.error("DNS服务器列表为空")
                            return False

                        logger.info(f"设置DNS服务器列表: {dns_servers}")
                        # 创建方法参数对象
                        method_params = (
                            target_config.Methods_("SetDNSServerSearchOrder")
                            .InParameters.SpawnInstance_()
                        )
                        method_params.DNSServerSearchOrder = list(dns_servers)
                        result = target_config.ExecMethod_(
                            "SetDNSServerSearchOrder", method_params
                        )

                    # 处理返回结果
                    error_code = self._extract_error_code(result)
                    if error_code == 0:
                        action = "DHCP DNS" if enable_dhcp else "静态DNS"
                        logger.info(f"WMI设置{action}成功")
                        return True
                    else:
                        action = "DHCP DNS" if enable_dhcp else "静态DNS"
                        logger.error(f"WMI设置{action}失败，错误代码: {error_code}")
                        return False

                except Exception as e:
                    logger.error(f"WMI方法调用失败: {e}")
                    raise

        except Exception as e:
            logger.error(f"设置DNS失败: {e}")
            return False

    def _extract_error_code(self, result) -> int:
        """
        从WMI结果中提取错误代码

        Args:
            result: WMI执行结果对象

        Returns:
            错误代码，0表示成功
        """
        try:
            if hasattr(result, "ReturnValue"):
                return result.ReturnValue
            elif hasattr(result, "Properties_"):
                try:
                    return_value_prop = result.Properties_("ReturnValue")
                    if return_value_prop:
                        return return_value_prop.Value
                except:
                    try:
                        for prop in result.Properties_:
                            if prop.Name == "ReturnValue":
                                return prop.Value
                    except:
                        pass
            try:
                return int(result)
            except:
                pass
        except Exception:
            pass
        return 0  # 假设成功

    def clear_dns_cache(self) -> bool:
        """
        清理系统DNS缓存

        Returns:
            bool: 清理成功返回True，失败返回False
        """
        if self.platform != "Windows":
            logger.error("此功能仅支持Windows系统")
            return False

        try:
            logger.info("开始清理DNS缓存")
            result = subprocess.run(
                ["ipconfig", "/flushdns"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if result.returncode == 0:
                logger.info("DNS缓存已清理")
                return True
            else:
                logger.error(f"清理DNS缓存失败: {result.stderr}")
                return False
        except Exception as e:
            logger.error(f"清理DNS缓存失败: {e}")
            return False
