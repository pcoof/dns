"""DNS测试器测试模块"""

import pytest
from unittest.mock import patch, MagicMock
from core.dns_tester import DNSTester, TestResult


class TestTestResult:
    """测试结果数据类测试"""

    def test_both_success(self):
        """测试主备DNS都成功"""
        result = TestResult(
            name="Test",
            primary_latency=10.0,
            primary_status="成功",
            secondary_latency=20.0,
            secondary_status="成功",
        )
        assert result.avg_latency == 15.0
        assert result.status == "成功"

    def test_primary_success_only(self):
        """测试只有主DNS成功"""
        result = TestResult(
            name="Test",
            primary_latency=10.0,
            primary_status="成功",
            secondary_latency=None,
            secondary_status=None,
        )
        assert result.avg_latency == 10.0
        assert result.status == "成功"

    def test_both_failed(self):
        """测试主备DNS都失败"""
        result = TestResult(
            name="Test",
            primary_latency=float("inf"),
            primary_status="失败",
            secondary_latency=float("inf"),
            secondary_status="失败",
        )
        assert result.avg_latency == float("inf")
        assert result.status == "失败"


class TestDNSTester:
    """DNS测试器测试"""

    def test_init(self):
        """测试初始化"""
        tester = DNSTester()
        assert tester.timeout == 3
        assert tester.max_workers == 5

    def test_init_custom_params(self):
        """测试自定义参数初始化"""
        tester = DNSTester(timeout=5, max_workers=10)
        assert tester.timeout == 5
        assert tester.max_workers == 10

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_success(self, mock_resolver):
        """测试DNS查询成功"""
        # Mock成功响应
        mock_answer = MagicMock()
        mock_resolver.return_value.resolve.return_value = [mock_answer]

        tester = DNSTester()
        latency, status = tester.test_dns("8.8.8.8", "www.baidu.com")

        assert status == "成功"
        assert latency > 0
        mock_resolver.return_value.resolve.assert_called_once()

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_timeout(self, mock_resolver):
        """测试DNS查询超时"""
        import dns.resolver
        mock_resolver.return_value.resolve.side_effect = dns.resolver.Timeout()

        tester = DNSTester()
        latency, status = tester.test_dns("8.8.8.8", "www.baidu.com")

        assert status == "失败"
        assert latency == float("inf")

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_invalid_ip(self, mock_resolver):
        """测试无效IP地址"""
        tester = DNSTester()
        latency, status = tester.test_dns("invalid", "www.baidu.com")

        assert status == "失败"
        assert latency == float("inf")
        mock_resolver.assert_not_called()

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_server_full(self, mock_resolver):
        """测试完整DNS服务器测试"""
        mock_answer = MagicMock()
        mock_resolver.return_value.resolve.return_value = [mock_answer]

        tester = DNSTester()
        result = tester.test_dns_server_full(
            name="Google",
            primary="8.8.8.8",
            secondary="8.8.4.4",
        )

        assert result.name == "Google"
        assert result.primary_status == "成功"
        assert result.secondary_status == "成功"

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_servers_parallel(self, mock_resolver):
        """测试并行DNS测试"""
        mock_answer = MagicMock()
        mock_resolver.return_value.resolve.return_value = [mock_answer]

        servers = [
            {"name": "Google", "primary": "8.8.8.8", "secondary": "8.8.4.4"},
            {"name": "Cloudflare", "primary": "1.1.1.1", "secondary": "1.0.0.1"},
        ]

        tester = DNSTester()
        results = tester.test_dns_servers_parallel(servers)

        assert len(results) == 2
        assert "Google" in results
        assert "Cloudflare" in results
        assert results["Google"].status == "成功"
        assert results["Cloudflare"].status == "成功"

    def test_test_dns_servers_empty(self):
        """测试空服务器列表"""
        tester = DNSTester()
        results = tester.test_dns_servers_parallel([])
        assert results == {}

    @patch("core.dns_tester.dns.resolver.Resolver")
    def test_test_dns_servers_sequential(self, mock_resolver):
        """测试顺序DNS测试"""
        mock_answer = MagicMock()
        mock_resolver.return_value.resolve.return_value = [mock_answer]

        servers = [
            {"name": "Google", "primary": "8.8.8.8", "secondary": "8.8.4.4"},
        ]

        tester = DNSTester()
        results = tester.test_dns_servers_sequential(servers)

        assert len(results) == 1
        assert "Google" in results