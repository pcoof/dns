"""导出/导入管理模块"""

import csv
import json
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime

from config.models import DNSServer
from utils.logger import get_logger

logger = get_logger("ExportManager")


class ExportManager:
    """导出/导入管理器"""

    def __init__(self):
        """初始化导出管理器"""
        logger.info("导出管理器初始化完成")

    def export_dns_to_csv(
        self,
        servers: List[DNSServer],
        test_results: Optional[Dict] = None,
        filename: Optional[str] = None,
    ) -> bool:
        """
        导出DNS服务器列表到CSV文件

        Args:
            servers: DNS服务器列表
            test_results: 测试结果字典（可选）
            filename: 输出文件名，如果为None则自动生成

        Returns:
            bool: 导出成功返回True，失败返回False
        """
        try:
            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"dns_export_{timestamp}.csv"

            filepath = Path(filename)
            filepath.parent.mkdir(parents=True, exist_ok=True)

            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)

                # 写入表头
                headers = ["名称", "主DNS", "备用DNS", "状态", "延迟(ms)"]
                writer.writerow(headers)

                # 写入数据
                for server in servers:
                    result = test_results.get(server.name, {}) if test_results else {}
                    status = result.get("status", "")
                    latency = result.get("avg_latency", "")

                    # 格式化延迟
                    if latency == float("inf"):
                        latency_str = "∞"
                    elif latency:
                        latency_str = f"{latency:.2f}"
                    else:
                        latency_str = ""

                    writer.writerow(
                        [server.name, server.primary, server.secondary, status, latency_str]
                    )

            logger.info(f"成功导出DNS列表到CSV: {filename}")
            return True

        except Exception as e:
            logger.error(f"导出CSV失败: {e}")
            return False

    def export_dns_to_json(
        self,
        servers: List[DNSServer],
        test_results: Optional[Dict] = None,
        filename: Optional[str] = None,
    ) -> bool:
        """
        导出DNS服务器列表到JSON文件

        Args:
            servers: DNS服务器列表
            test_results: 测试结果字典（可选）
            filename: 输出文件名，如果为None则自动生成

        Returns:
            bool: 导出成功返回True，失败返回False
        """
        try:
            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"dns_export_{timestamp}.json"

            filepath = Path(filename)
            filepath.parent.mkdir(parents=True, exist_ok=True)

            # 准备导出数据
            export_data = {
                "export_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "total_servers": len(servers),
                "servers": [],
            }

            for server in servers:
                server_data = {
                    "name": server.name,
                    "primary": server.primary,
                    "secondary": server.secondary,
                    "enabled": server.enabled,
                    "tags": server.tags,
                }

                # 添加测试结果
                if test_results and server.name in test_results:
                    result = test_results[server.name]
                    server_data["test_result"] = {
                        "status": result.status,
                        "avg_latency": result.avg_latency,
                        "primary_latency": result.primary_latency,
                        "primary_status": result.primary_status,
                        "secondary_latency": result.secondary_latency,
                        "secondary_status": result.secondary_status,
                    }

                export_data["servers"].append(server_data)

            # 写入JSON文件
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)

            logger.info(f"成功导出DNS列表到JSON: {filename}")
            return True

        except Exception as e:
            logger.error(f"导出JSON失败: {e}")
            return False

    def import_dns_from_json(self, filename: str) -> Optional[List[DNSServer]]:
        """
        从JSON文件导入DNS服务器列表

        Args:
            filename: JSON文件路径

        Returns:
            DNS服务器列表，导入失败返回None
        """
        try:
            filepath = Path(filename)
            if not filepath.exists():
                logger.error(f"文件不存在: {filename}")
                return None

            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)

            servers = []
            for server_data in data.get("servers", []):
                server = DNSServer(
                    name=server_data.get("name", ""),
                    primary=server_data.get("primary", ""),
                    secondary=server_data.get("secondary", ""),
                    enabled=server_data.get("enabled", True),
                    tags=server_data.get("tags", []),
                )
                servers.append(server)

            logger.info(f"成功从JSON导入 {len(servers)} 个DNS服务器")
            return servers

        except json.JSONDecodeError as e:
            logger.error(f"JSON格式错误: {e}")
            return None
        except Exception as e:
            logger.error(f"导入JSON失败: {e}")
            return None

    def import_dns_from_csv(self, filename: str) -> Optional[List[DNSServer]]:
        """
        从CSV文件导入DNS服务器列表

        Args:
            filename: CSV文件路径

        Returns:
            DNS服务器列表，导入失败返回None
        """
        try:
            filepath = Path(filename)
            if not filepath.exists():
                logger.error(f"文件不存在: {filename}")
                return None

            servers = []
            with open(filepath, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)

                for row in reader:
                    server = DNSServer(
                        name=row.get("名称", "").strip(),
                        primary=row.get("主DNS", "").strip(),
                        secondary=row.get("备用DNS", "").strip(),
                    )
                    servers.append(server)

            logger.info(f"成功从CSV导入 {len(servers)} 个DNS服务器")
            return servers

        except Exception as e:
            logger.error(f"导入CSV失败: {e}")
            return None

    def export_history_to_json(
        self,
        history_records: List,
        filename: Optional[str] = None,
    ) -> bool:
        """
        导出历史记录到JSON文件

        Args:
            history_records: 历史记录列表
            filename: 输出文件名

        Returns:
            bool: 导出成功返回True，失败返回False
        """
        try:
            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"dns_history_{timestamp}.json"

            filepath = Path(filename)
            filepath.parent.mkdir(parents=True, exist_ok=True)

            export_data = {
                "export_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "total_records": len(history_records),
                "records": [record.to_dict() for record in history_records],
            }

            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)

            logger.info(f"成功导出历史记录到JSON: {filename}")
            return True

        except Exception as e:
            logger.error(f"导出历史记录失败: {e}")
            return False