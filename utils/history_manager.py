"""历史记录管理模块"""

import json
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime
from dataclasses import dataclass, asdict

from config.settings import AppConfig
from utils.logger import get_logger

logger = get_logger("HistoryManager")


@dataclass
class HistoryRecord:
    """历史记录数据类"""

    dns_name: str
    dns_primary: str
    dns_secondary: str
    latency: float
    status: str
    timestamp: str
    category: str = ""

    def to_dict(self) -> Dict:
        """转换为字典"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict) -> "HistoryRecord":
        """从字典创建实例"""
        return cls(**data)


class HistoryManager:
    """DNS测试历史记录管理器"""

    def __init__(self, history_file: Optional[str] = None, max_records: int = 1000):
        """
        初始化历史记录管理器

        Args:
            history_file: 历史记录文件路径
            max_records: 最大记录数
        """
        self.history_file = history_file or AppConfig.HISTORY_FILE
        self.max_records = max_records
        self.history: List[HistoryRecord] = []
        self._load_history()
        logger.info(f"历史记录管理器初始化完成，文件: {self.history_file}")

    def _load_history(self) -> None:
        """加载历史记录"""
        try:
            history_path = Path(self.history_file)
            if not history_path.exists():
                logger.info("历史记录文件不存在，将创建新文件")
                return

            with open(self.history_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.history = [HistoryRecord.from_dict(item) for item in data]
            logger.info(f"成功加载 {len(self.history)} 条历史记录")

        except json.JSONDecodeError as e:
            logger.error(f"历史记录文件格式错误: {e}")
            self.history = []
        except Exception as e:
            logger.error(f"加载历史记录失败: {e}")
            self.history = []

    def _save_history(self) -> None:
        """保存历史记录"""
        try:
            history_path = Path(self.history_file)
            history_path.parent.mkdir(parents=True, exist_ok=True)

            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(
                    [record.to_dict() for record in self.history],
                    f,
                    ensure_ascii=False,
                    indent=2,
                )

            logger.info(f"成功保存 {len(self.history)} 条历史记录")

        except Exception as e:
            logger.error(f"保存历史记录失败: {e}")

    def add_record(
        self,
        dns_name: str,
        dns_primary: str,
        dns_secondary: str,
        latency: float,
        status: str,
        category: str = "",
    ) -> None:
        """
        添加历史记录

        Args:
            dns_name: DNS服务器名称
            dns_primary: 主DNS地址
            dns_secondary: 备用DNS地址
            latency: 延迟（毫秒）
            status: 状态（成功/失败）
            category: 分类名称
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = HistoryRecord(
            dns_name=dns_name,
            dns_primary=dns_primary,
            dns_secondary=dns_secondary,
            latency=latency,
            status=status,
            timestamp=timestamp,
            category=category,
        )

        self.history.append(record)

        # 限制记录数量
        if len(self.history) > self.max_records:
            self.history = self.history[-self.max_records :]

        self._save_history()
        logger.info(f"添加历史记录: {dns_name} - {status}")

    def get_records(
        self,
        limit: Optional[int] = None,
        category: Optional[str] = None,
        status: Optional[str] = None,
    ) -> List[HistoryRecord]:
        """
        获取历史记录

        Args:
            limit: 返回记录数量限制
            category: 按分类过滤
            status: 按状态过滤

        Returns:
            历史记录列表
        """
        records = self.history.copy()

        # 按分类过滤
        if category:
            records = [r for r in records if r.category == category]

        # 按状态过滤
        if status:
            records = [r for r in records if r.status == status]

        # 限制数量
        if limit:
            records = records[-limit:]

        return records

    def get_records_by_dns(self, dns_name: str, limit: int = 10) -> List[HistoryRecord]:
        """
        获取指定DNS服务器的历史记录

        Args:
            dns_name: DNS服务器名称
            limit: 返回记录数量限制

        Returns:
            历史记录列表
        """
        records = [r for r in self.history if r.dns_name == dns_name]
        return records[-limit:] if limit else records

    def clear_history(self) -> None:
        """清空历史记录"""
        self.history = []
        self._save_history()
        logger.info("已清空历史记录")

    def get_statistics(self) -> Dict:
        """
        获取历史记录统计信息

        Returns:
            统计信息字典
        """
        if not self.history:
            return {
                "total": 0,
                "success": 0,
                "failed": 0,
                "categories": {},
                "avg_latency": 0,
            }

        total = len(self.history)
        success = sum(1 for r in self.history if r.status == "成功")
        failed = total - success

        # 按分类统计
        categories = {}
        for record in self.history:
            cat = record.category or "未分类"
            if cat not in categories:
                categories[cat] = {"total": 0, "success": 0, "failed": 0}
            categories[cat]["total"] += 1
            if record.status == "成功":
                categories[cat]["success"] += 1
            else:
                categories[cat]["failed"] += 1

        # 计算平均延迟（只计算成功的）
        successful_records = [r for r in self.history if r.status == "成功"]
        avg_latency = (
            sum(r.latency for r in successful_records) / len(successful_records)
            if successful_records
            else 0
        )

        return {
            "total": total,
            "success": success,
            "failed": failed,
            "categories": categories,
            "avg_latency": avg_latency,
        }