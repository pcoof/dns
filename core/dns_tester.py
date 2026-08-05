"""DNS测试器模块 - 支持并行测试"""

import dns.resolver
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Tuple, Optional, Callable
from dataclasses import dataclass

from config.settings import AppConfig
from utils.logger import get_logger

logger = get_logger("DNSTester")


@dataclass
class TestResult:
    """DNS测试结果数据类"""

    name: str
    primary_latency: float
    primary_status: str
    secondary_latency: Optional[float] = None
    secondary_status: Optional[str] = None
    avg_latency: float = 0.0
    status: str = ""

    def __post_init__(self):
        """计算平均延迟和总体状态"""
        if self.primary_status == "成功" and self.secondary_status == "成功":
            self.avg_latency = (self.primary_latency + self.secondary_latency) / 2
            self.status = "成功"
        elif self.primary_status == "成功":
            self.avg_latency = self.primary_latency
            self.status = "成功"
        elif self.secondary_status == "成功":
            self.avg_latency = self.secondary_latency
            self.status = "成功"
        else:
            self.avg_latency = float("inf")
            self.status = "失败"


class DNSTester:
    """DNS服务器测试器"""

    def __init__(
        self,
        timeout: int = None,
        max_workers: int = None,
        progress_callback: Optional[Callable[[str, int, int], None]] = None,
    ):
        """
        初始化DNS测试器

        Args:
            timeout: DNS查询超时时间（秒），默认使用AppConfig中的配置
            max_workers: 最大并发线程数，默认使用AppConfig中的配置
            progress_callback: 进度回调函数，参数为(服务器名称, 当前索引, 总数)
        """
        self.timeout = timeout or AppConfig.DNS_TEST_TIMEOUT
        self.max_workers = max_workers or AppConfig.DNS_TEST_MAX_WORKERS
        self.progress_callback = progress_callback
        logger.info(
            f"DNS测试器初始化完成，超时: {self.timeout}s, 并发数: {self.max_workers}"
        )

    def test_dns(self, dns_server: str, domain: str = None) -> Tuple[float, str]:
        """
        测试单个DNS服务器的响应时间和可用性

        Args:
            dns_server: DNS服务器IP地址
            domain: 要解析的域名，默认使用AppConfig中的配置

        Returns:
            tuple: (延迟毫秒数, 状态字符串)
                - 延迟: 成功时为毫秒数，失败时为float('inf')
                - 状态: "成功" 或 "失败"

        Raises:
            ValueError: 当DNS服务器地址无效时

        Example:
            >>> tester = DNSTester()
            >>> latency, status = tester.test_dns("8.8.8.8", "www.baidu.com")
            >>> print(f"延迟: {latency}ms, 状态: {status}")
        """
        from utils.helpers import validate_dns_ip

        if not validate_dns_ip(dns_server):
            logger.error(f"无效的DNS服务器地址: {dns_server}")
            return float("inf"), "失败"

        domain = domain or AppConfig.DNS_TEST_DOMAIN

        try:
            resolver = dns.resolver.Resolver()
            resolver.nameservers = [dns_server]
            resolver.timeout = self.timeout
            resolver.lifetime = self.timeout

            start_time = time.time()
            answers = resolver.resolve(domain)
            end_time = time.time()

            latency = (end_time - start_time) * 1000
            logger.info(f"DNS测试成功: {dns_server} - 延迟: {latency:.2f}ms")
            return latency, "成功"

        except dns.resolver.Timeout:
            logger.warning(f"DNS查询超时: {dns_server}")
            return float("inf"), "失败"
        except dns.resolver.NoAnswer:
            logger.warning(f"DNS服务器无响应: {dns_server}")
            return float("inf"), "失败"
        except dns.resolver.NXDOMAIN:
            logger.warning(f"域名不存在: {domain} (DNS: {dns_server})")
            return float("inf"), "失败"
        except dns.resolver.NoNameservers:
            logger.warning(f"无可用的DNS服务器: {dns_server}")
            return float("inf"), "失败"
        except Exception as e:
            logger.error(f"DNS测试失败: {dns_server} - 错误: {e}")
            return float("inf"), "失败"

    def test_dns_server_full(
        self, name: str, primary: str, secondary: str = ""
    ) -> TestResult:
        """
        完整测试DNS服务器（主DNS和备用DNS）

        Args:
            name: DNS服务器名称
            primary: 主DNS地址
            secondary: 备用DNS地址

        Returns:
            TestResult: 测试结果对象
        """
        # 测试主DNS
        primary_latency, primary_status = self.test_dns(primary)

        # 测试备用DNS（如果主DNS失败且存在备用DNS）
        secondary_latency = None
        secondary_status = None

        if primary_status == "失败" and secondary:
            secondary_latency, secondary_status = self.test_dns(secondary)
        elif secondary:
            # 主DNS成功，也测试备用DNS
            secondary_latency, secondary_status = self.test_dns(secondary)

        return TestResult(
            name=name,
            primary_latency=primary_latency,
            primary_status=primary_status,
            secondary_latency=secondary_latency,
            secondary_status=secondary_status,
        )

    def test_dns_servers_parallel(
        self,
        servers: List[Dict[str, str]],
        domain: str = None,
    ) -> Dict[str, TestResult]:
        """
        并行测试多个DNS服务器

        Args:
            servers: DNS服务器列表，每个元素为包含name, primary, secondary的字典
            domain: 要解析的域名，默认使用AppConfig中的配置

        Returns:
            dict: 服务器名称到测试结果的映射

        Example:
            >>> servers = [
            ...     {"name": "Google", "primary": "8.8.8.8", "secondary": "8.8.4.4"},
            ...     {"name": "Cloudflare", "primary": "1.1.1.1", "secondary": "1.0.0.1"},
            ... ]
            >>> tester = DNSTester()
            >>> results = tester.test_dns_servers_parallel(servers)
            >>> for name, result in results.items():
            ...     print(f"{name}: {result.status} - {result.avg_latency}ms")
        """
        if not servers:
            logger.warning("没有DNS服务器需要测试")
            return {}

        results = {}
        total = len(servers)
        domain = domain or AppConfig.DNS_TEST_DOMAIN

        logger.info(f"开始并行测试 {total} 个DNS服务器")

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # 提交所有测试任务
            future_to_server = {
                executor.submit(
                    self.test_dns_server_full,
                    server["name"],
                    server["primary"],
                    server.get("secondary", ""),
                ): server
                for server in servers
            }

            # 处理完成的任务
            for i, future in enumerate(as_completed(future_to_server), 1):
                server = future_to_server[future]
                try:
                    result = future.result()
                    results[result.name] = result

                    # 调用进度回调
                    if self.progress_callback:
                        self.progress_callback(server["name"], i, total)

                except Exception as e:
                    logger.error(
                        f"测试 {server['name']} 时发生异常: {e}"
                    )
                    # 创建失败结果
                    results[server["name"]] = TestResult(
                        name=server["name"],
                        primary_latency=float("inf"),
                        primary_status="失败",
                        secondary_status="失败",
                    )

        logger.info(f"并行测试完成，共测试 {len(results)} 个服务器")
        return results

    def test_dns_servers_sequential(
        self,
        servers: List[Dict[str, str]],
        domain: str = None,
    ) -> Dict[str, TestResult]:
        """
        顺序测试多个DNS服务器（兼容旧版本）

        Args:
            servers: DNS服务器列表
            domain: 要解析的域名

        Returns:
            dict: 服务器名称到测试结果的映射
        """
        if not servers:
            logger.warning("没有DNS服务器需要测试")
            return {}

        results = {}
        total = len(servers)
        domain = domain or AppConfig.DNS_TEST_DOMAIN

        logger.info(f"开始顺序测试 {total} 个DNS服务器")

        for i, server in enumerate(servers, 1):
            try:
                result = self.test_dns_server_full(
                    server["name"],
                    server["primary"],
                    server.get("secondary", ""),
                )
                results[result.name] = result

                # 调用进度回调
                if self.progress_callback:
                    self.progress_callback(server["name"], i, total)

            except Exception as e:
                logger.error(f"测试 {server['name']} 时发生异常: {e}")
                results[server["name"]] = TestResult(
                    name=server["name"],
                    primary_latency=float("inf"),
                    primary_status="失败",
                    secondary_status="失败",
                )

        logger.info(f"顺序测试完成，共测试 {len(results)} 个服务器")
        return results