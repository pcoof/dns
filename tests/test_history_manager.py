"""历史记录管理器测试模块"""

import pytest
import os
import tempfile
from utils.history_manager import HistoryManager, HistoryRecord


class TestHistoryRecord:
    """历史记录数据类测试"""

    def test_to_dict(self):
        """测试转换为字典"""
        record = HistoryRecord(
            dns_name="Google",
            dns_primary="8.8.8.8",
            dns_secondary="8.8.4.4",
            latency=10.5,
            status="成功",
            timestamp="2024-01-01 12:00:00",
            category="Ipv4_默认",
        )
        data = record.to_dict()
        assert data["dns_name"] == "Google"
        assert data["latency"] == 10.5
        assert data["status"] == "成功"

    def test_from_dict(self):
        """测试从字典创建"""
        data = {
            "dns_name": "Cloudflare",
            "dns_primary": "1.1.1.1",
            "dns_secondary": "1.0.0.1",
            "latency": 5.2,
            "status": "成功",
            "timestamp": "2024-01-01 12:00:00",
            "category": "Ipv4_默认",
        }
        record = HistoryRecord.from_dict(data)
        assert record.dns_name == "Cloudflare"
        assert record.latency == 5.2


class TestHistoryManager:
    """历史记录管理器测试"""

    @pytest.fixture
    def temp_history_file(self):
        """创建临时历史文件"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            f.write("[]\n")
            temp_file = f.name
        yield temp_file
        # 清理
        if os.path.exists(temp_file):
            os.remove(temp_file)

    def test_init(self, temp_history_file):
        """测试初始化"""
        manager = HistoryManager(history_file=temp_history_file)
        assert manager.history_file == temp_history_file
        assert manager.max_records == 1000

    def test_init_custom_max_records(self, temp_history_file):
        """测试自定义最大记录数"""
        manager = HistoryManager(history_file=temp_history_file, max_records=100)
        assert manager.max_records == 100

    def test_add_record(self, temp_history_file):
        """测试添加记录"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record(
            dns_name="Google",
            dns_primary="8.8.8.8",
            dns_secondary="8.8.4.4",
            latency=10.5,
            status="成功",
            category="Ipv4_默认",
        )
        assert len(manager.history) == 1
        assert manager.history[0].dns_name == "Google"

    def test_get_records(self, temp_history_file):
        """测试获取记录"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        manager.add_record("Cloudflare", "1.1.1.1", "1.0.0.1", 5.2, "成功", "Ipv4_默认")

        records = manager.get_records()
        assert len(records) == 2

    def test_get_records_with_limit(self, temp_history_file):
        """测试获取记录（带限制）"""
        manager = HistoryManager(history_file=temp_history_file)
        for i in range(10):
            manager.add_record(f"Server{i}", "8.8.8.8", "8.8.4.4", 10.0, "成功", "Ipv4_默认")

        records = manager.get_records(limit=5)
        assert len(records) == 5

    def test_get_records_by_category(self, temp_history_file):
        """测试按分类获取记录"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        manager.add_record("Cloudflare", "1.1.1.1", "1.0.0.1", 5.2, "成功", "Ipv4_家庭")

        records = manager.get_records(category="Ipv4_默认")
        assert len(records) == 1
        assert records[0].dns_name == "Google"

    def test_get_records_by_status(self, temp_history_file):
        """测试按状态获取记录"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        manager.add_record("Cloudflare", "1.1.1.1", "1.0.0.1", float("inf"), "失败", "Ipv4_默认")

        records = manager.get_records(status="成功")
        assert len(records) == 1
        assert records[0].dns_name == "Google"

    def test_get_records_by_dns(self, temp_history_file):
        """测试按DNS名称获取记录"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 12.3, "成功", "Ipv4_默认")
        manager.add_record("Cloudflare", "1.1.1.1", "1.0.0.1", 5.2, "成功", "Ipv4_默认")

        records = manager.get_records_by_dns("Google")
        assert len(records) == 2

    def test_clear_history(self, temp_history_file):
        """测试清空历史"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        assert len(manager.history) == 1

        manager.clear_history()
        assert len(manager.history) == 0

    def test_get_statistics(self, temp_history_file):
        """测试获取统计信息"""
        manager = HistoryManager(history_file=temp_history_file)
        manager.add_record("Google", "8.8.8.8", "8.8.4.4", 10.5, "成功", "Ipv4_默认")
        manager.add_record("Cloudflare", "1.1.1.1", "1.0.0.1", 5.2, "成功", "Ipv4_家庭")
        manager.add_record("Invalid", "9.9.9.9", "9.9.9.10", float("inf"), "失败", "Ipv4_默认")

        stats = manager.get_statistics()
        assert stats["total"] == 3
        assert stats["success"] == 2
        assert stats["failed"] == 1
        assert "Ipv4_默认" in stats["categories"]
        assert "Ipv4_家庭" in stats["categories"]
        assert stats["avg_latency"] == (10.5 + 5.2) / 2

    def test_get_statistics_empty(self, temp_history_file):
        """测试空历史统计"""
        manager = HistoryManager(history_file=temp_history_file)
        stats = manager.get_statistics()
        assert stats["total"] == 0
        assert stats["success"] == 0
        assert stats["failed"] == 0

    def test_max_records_limit(self, temp_history_file):
        """测试最大记录数限制"""
        manager = HistoryManager(history_file=temp_history_file, max_records=5)
        for i in range(10):
            manager.add_record(f"Server{i}", "8.8.8.8", "8.8.4.4", 10.0, "成功", "Ipv4_默认")

        assert len(manager.history) == 5
        # 应该保留最后5条记录
        assert manager.history[0].dns_name == "Server5"
        assert manager.history[-1].dns_name == "Server9"