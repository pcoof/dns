"""数据模型模块"""

from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any


@dataclass
class DNSServer:
    """DNS服务器数据类"""

    name: str
    primary: str
    secondary: str = ""
    enabled: bool = True
    tags: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DNSServer":
        """从字典创建实例"""
        return cls(
            name=data.get("name", ""),
            primary=data.get("primary", ""),
            secondary=data.get("secondary", ""),
            enabled=data.get("enabled", True),
            tags=data.get("tags", []),
        )


@dataclass
class Category:
    """DNS分类数据类"""

    name: str
    servers: List[DNSServer] = field(default_factory=list)
    order: int = 0

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "name": self.name,
            "servers": [server.to_dict() for server in self.servers],
            "order": self.order,
        }