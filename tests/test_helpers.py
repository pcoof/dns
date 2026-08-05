"""辅助函数测试模块"""

import pytest
from utils.helpers import (
    validate_dns_ip,
    sanitize_input,
    is_ip,
    get_notification_style,
)
from ttkbootstrap.constants import SUCCESS, DANGER, WARNING, INFO


class TestValidateDNSIP:
    """DNS IP验证测试"""

    def test_valid_ipv4(self):
        """测试有效的IPv4地址"""
        assert validate_dns_ip("8.8.8.8") is True
        assert validate_dns_ip("1.1.1.1") is True
        assert validate_dns_ip("223.5.5.5") is True

    def test_invalid_ipv4(self):
        """测试无效的IPv4地址"""
        assert validate_dns_ip("256.256.256.256") is False
        assert validate_dns_ip("invalid") is False
        assert validate_dns_ip("8.8.8") is False
        assert validate_dns_ip("") is False

    def test_valid_ipv6(self):
        """测试有效的IPv6地址"""
        assert validate_dns_ip("2001:4860:4860::8888") is True
        assert validate_dns_ip("2606:4700:4700::1111") is True

    def test_invalid_ipv6(self):
        """测试无效的IPv6地址"""
        assert validate_dns_ip("invalid") is False
        assert validate_dns_ip("") is False


class TestSanitizeInput:
    """输入清理测试"""

    def test_normal_input(self):
        """测试正常输入"""
        assert sanitize_input("test") == "test"
        assert sanitize_input("  test  ") == "test"

    def test_html_tags(self):
        """测试HTML标签清理"""
        assert sanitize_input("<script>alert('xss')</script>") == "scriptalert(xss)/script"
        assert sanitize_input("<b>bold</b>") == "bbold/b"

    def test_special_characters(self):
        """测试特殊字符清理"""
        assert sanitize_input('test"quote') == "testquote"
        assert sanitize_input("test'quote") == "testquote"
        assert sanitize_input("test&symbol") == "testsymbol"


class TestIsIP:
    """IP地址检查测试"""

    def test_ipv4_only(self):
        """测试只检查IPv4"""
        assert is_ip("8.8.8.8", version=4) is True
        assert is_ip("2001:4860:4860::8888", version=4) is False

    def test_ipv6_only(self):
        """测试只检查IPv6"""
        assert is_ip("2001:4860:4860::8888", version=6) is True
        assert is_ip("8.8.8.8", version=6) is False

    def test_any_version(self):
        """测试检查任意版本"""
        assert is_ip("8.8.8.8") is True
        assert is_ip("2001:4860:4860::8888") is True
        assert is_ip("invalid") is False


class TestGetNotificationStyle:
    """通知样式测试"""

    def test_success_style(self):
        """测试成功样式"""
        prefix, color = get_notification_style(SUCCESS)
        assert prefix == "✓ "
        assert color == "green"

    def test_danger_style(self):
        """测试危险样式"""
        prefix, color = get_notification_style(DANGER)
        assert prefix == "✗ "
        assert color == "red"

    def test_warning_style(self):
        """测试警告样式"""
        prefix, color = get_notification_style(WARNING)
        assert prefix == "⚠ "
        assert color == "orange"

    def test_info_style(self):
        """测试信息样式"""
        prefix, color = get_notification_style(INFO)
        assert prefix == "ℹ "
        assert color == "blue"

    def test_invalid_style(self):
        """测试无效样式"""
        prefix, color = get_notification_style("invalid")
        assert prefix == "ℹ "
        assert color == "blue"