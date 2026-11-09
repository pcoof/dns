"""配置管理器测试模块"""

import pytest
import os
import tempfile

from config.config_manager import ConfigManager, DNSServer


class TestDNSServer:
    """DNS服务器数据类测试"""

    def test_to_dict(self):
        """测试转换为字典"""
        server = DNSServer(
            name="Google",
            primary="8.8.8.8",
            secondary="8.8.4.4",
            enabled=True,
            tags=["fast", "reliable"],
        )
        data = server.to_dict()
        assert data["name"] == "Google"
        assert data["primary"] == "8.8.8.8"
        assert data["secondary"] == "8.8.4.4"
        assert data["enabled"] is True
        assert data["tags"] == ["fast", "reliable"]

    def test_from_dict(self):
        """测试从字典创建"""
        data = {
            "name": "Cloudflare",
            "primary": "1.1.1.1",
            "secondary": "1.0.0.1",
            "enabled": True,
            "tags": ["secure"],
        }
        server = DNSServer.from_dict(data)
        assert server.name == "Cloudflare"
        assert server.primary == "1.1.1.1"
        assert server.secondary == "1.0.0.1"


class TestConfigManager:
    """配置管理器测试"""

    @pytest.fixture
    def temp_config_file(self):
        """创建临时配置文件"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
            f.write("[TestCategory]\n")
            f.write("TestServer = 8.8.8.8,8.8.4.4\n")
            temp_file = f.name
        yield temp_file
        # 清理
        if os.path.exists(temp_file):
            os.remove(temp_file)

    def test_init(self, temp_config_file):
        """测试初始化"""
        manager = ConfigManager(config_file=temp_config_file)
        assert manager.config_file == temp_config_file

    def test_load_config(self, temp_config_file):
        """测试加载配置"""
        manager = ConfigManager(config_file=temp_config_file)
        result = manager.load_config()
        assert result is True
        assert "TestCategory" in manager.config.sections()

    def test_get_categories(self, temp_config_file):
        """测试获取分类"""
        manager = ConfigManager(config_file=temp_config_file)
        manager.load_config()
        categories = manager.get_categories()
        assert "TestCategory" in categories

    def test_save_config(self):
        """测试保存配置"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
            temp_file = f.name

        try:
            manager = ConfigManager(config_file=temp_file)
            manager.config.add_section("Test")
            manager.config.set("Test", "key", "value")
            result = manager.save_config()
            assert result is True
            assert os.path.exists(temp_file)
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_get_category_dns_servers(self, temp_config_file):
        """测试获取分类DNS服务器"""
        manager = ConfigManager(config_file=temp_config_file)
        manager.load_config()
        servers = manager.get_category_dns_servers("TestCategory")
        assert len(servers) == 1
        assert servers[0].name == "TestServer"
        assert servers[0].primary == "8.8.8.8"
        assert servers[0].secondary == "8.8.4.4"

    def test_save_category_dns_servers(self):
        """测试保存分类DNS服务器"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
            temp_file = f.name

        try:
            manager = ConfigManager(config_file=temp_file)
            servers = [
                DNSServer(name="Google", primary="8.8.8.8", secondary="8.8.4.4"),
                DNSServer(name="Cloudflare", primary="1.1.1.1", secondary="1.0.0.1"),
            ]
            result = manager.save_category_dns_servers("TestCategory", servers)
            assert result is True

            # 验证保存
            manager.load_config()
            loaded_servers = manager.get_category_dns_servers("TestCategory")
            assert len(loaded_servers) == 2
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_add_category(self):
        """测试添加分类"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
            temp_file = f.name

        try:
            manager = ConfigManager(config_file=temp_file)
            manager.load_config()
            result = manager.add_category("NewCategory")
            assert result is True
            assert "NewCategory" in manager.config.sections()
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)

    def test_delete_category(self):
        """测试删除分类"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
            f.write("[ToDelete]\n")
            temp_file = f.name

        try:
            manager = ConfigManager(config_file=temp_file)
            manager.load_config()
            result = manager.delete_category("ToDelete")
            assert result is True
            assert "ToDelete" not in manager.config.sections()
        finally:
            if os.path.exists(temp_file):
                os.remove(temp_file)