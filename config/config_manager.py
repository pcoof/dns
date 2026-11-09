"""配置文件管理模块"""

import configparser
from pathlib import Path
from typing import List, Optional
from .settings import AppConfig
from .models import DNSServer
from utils.logger import get_logger

logger = get_logger("ConfigManager")


class ConfigManager:
    """配置文件管理器"""

    def __init__(self, config_file: Optional[str] = None):
        """
        初始化配置管理器

        Args:
            config_file: 配置文件路径，默认使用AppConfig中的配置
        """
        self.config_file = config_file or AppConfig.CONFIG_FILE
        self.config = configparser.ConfigParser()
        self.config.optionxform = str  # 区分大小写
        logger.info(f"配置管理器初始化，配置文件: {self.config_file}")

    def load_config(self) -> bool:
        """
        加载配置文件

        Returns:
            bool: 加载成功返回True，失败返回False
        """
        try:
            config_path = Path(self.config_file)
            if not config_path.exists():
                logger.warning(f"配置文件不存在: {self.config_file}")
                return False

            self.config.read(self.config_file, encoding="utf-8")
            logger.info(f"成功加载配置文件: {self.config_file}")
            return True
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}")
            return False

    def save_config(self) -> bool:
        """
        保存配置文件

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            config_path = Path(self.config_file)
            config_path.parent.mkdir(parents=True, exist_ok=True)

            with open(self.config_file, "w", encoding="utf-8") as f:
                self.config.write(f)

            logger.info(f"成功保存配置文件: {self.config_file}")
            return True
        except PermissionError as e:
            logger.error(f"保存配置文件失败（权限不足）: {e}")
            return False
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
            return False

    def get_categories(self) -> List[str]:
        """
        获取所有分类名称

        Returns:
            分类名称列表
        """
        return [
            section
            for section in self.config.sections()
            if section != AppConfig.RESERVED_SECTION_NAME
        ]

    def get_category_order(self) -> List[str]:
        """
        获取分类顺序

        Returns:
            排序后的分类名称列表
        """
        categories = self.get_categories()
        if AppConfig.RESERVED_SECTION_NAME in self.config.sections():
            if self.config.has_option(
                AppConfig.RESERVED_SECTION_NAME, "category_order"
            ):
                saved_order = self.config.get(
                    AppConfig.RESERVED_SECTION_NAME, "category_order"
                ).split(",")
                # 保留保存的顺序，添加新分类
                ordered = [cat for cat in saved_order if cat in categories]
                for cat in categories:
                    if cat not in ordered:
                        ordered.append(cat)
                return ordered
        return categories

    def set_category_order(self, categories: List[str]) -> bool:
        """
        设置分类顺序

        Args:
            categories: 排序后的分类名称列表

        Returns:
            bool: 设置成功返回True，失败返回False
        """
        try:
            if AppConfig.RESERVED_SECTION_NAME not in self.config.sections():
                self.config.add_section(AppConfig.RESERVED_SECTION_NAME)

            self.config.set(
                AppConfig.RESERVED_SECTION_NAME,
                "category_order",
                ",".join(categories),
            )
            # 保存配置
            return self.save_config()
        except Exception as e:
            logger.error(f"设置分类顺序失败: {e}")
            return False

    def get_last_category(self) -> Optional[str]:
        """
        获取上次选择的分类

        Returns:
            上次选择的分类名称，如果不存在返回None
        """
        if AppConfig.RESERVED_SECTION_NAME in self.config.sections():
            if self.config.has_option(
                AppConfig.RESERVED_SECTION_NAME, "last_category"
            ):
                return self.config.get(
                    AppConfig.RESERVED_SECTION_NAME, "last_category"
                )
        return None

    def set_last_category(self, category: str) -> bool:
        """
        设置上次选择的分类

        Args:
            category: 分类名称

        Returns:
            bool: 设置成功返回True，失败返回False
        """
        try:
            if AppConfig.RESERVED_SECTION_NAME not in self.config.sections():
                self.config.add_section(AppConfig.RESERVED_SECTION_NAME)

            self.config.set(
                AppConfig.RESERVED_SECTION_NAME, "last_category", category
            )
            # 保存配置
            return self.save_config()
        except Exception as e:
            logger.error(f"设置上次选择的分类失败: {e}")
            return False

    def get_theme(self) -> str:
        """
        获取保存的主题

        Returns:
            主题名称，如果不存在返回默认主题
        """
        if AppConfig.RESERVED_SECTION_NAME in self.config.sections():
            if self.config.has_option(AppConfig.RESERVED_SECTION_NAME, "theme"):
                return self.config.get(AppConfig.RESERVED_SECTION_NAME, "theme")
        return "darkly"

    def set_theme(self, theme: str) -> bool:
        """
        设置主题

        Args:
            theme: 主题名称

        Returns:
            bool: 设置成功返回True，失败返回False
        """
        try:
            if AppConfig.RESERVED_SECTION_NAME not in self.config.sections():
                self.config.add_section(AppConfig.RESERVED_SECTION_NAME)

            self.config.set(AppConfig.RESERVED_SECTION_NAME, "theme", theme.lower())
            # 保存配置
            return self.save_config()
        except Exception as e:
            logger.error(f"设置主题失败: {e}")
            return False

    def get_category_dns_servers(self, category: str) -> List[DNSServer]:
        """
        获取指定分类的DNS服务器列表

        Args:
            category: 分类名称

        Returns:
            DNS服务器列表
        """
        servers = []
        if category not in self.config.sections():
            logger.warning(f"分类不存在: {category}")
            return servers

        for key, value in self.config.items(category):
            if not value.strip():
                continue

            # 解析格式: name=primary,secondary[,other]
            dns_parts = value.split(",")
            if len(dns_parts) >= 2:
                primary = dns_parts[0].strip()
                secondary = dns_parts[1].strip()
            elif len(dns_parts) == 1:
                primary = dns_parts[0].strip()
                secondary = ""
            else:
                continue

            # 验证IP地址（接受合法的IPv4或IPv6）
            if primary and not self._is_valid_ip(primary):
                logger.warning(f"跳过无效IP: {key} - {primary}")
                continue

            if secondary and not self._is_valid_ip(secondary):
                secondary = ""

            if primary:
                servers.append(
                    DNSServer(name=key, primary=primary, secondary=secondary)
                )

        logger.info(f"从分类 {category} 加载了 {len(servers)} 个DNS服务器")
        return servers

    def _is_valid_ip(self, ip: str) -> bool:
        """
        验证IP地址是否合法（IPv4或IPv6均可）

        Args:
            ip: IP地址

        Returns:
            bool: 有效返回True，无效返回False
        """
        from utils.helpers import is_ip
        return is_ip(ip, version=4) or is_ip(ip, version=6)

    def save_category_dns_servers(
        self, category: str, servers: List[DNSServer]
    ) -> bool:
        """
        保存分类的DNS服务器列表

        Args:
            category: 分类名称
            servers: DNS服务器列表

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            # 清空该分类
            if category in self.config.sections():
                self.config.remove_section(category)
            self.config.add_section(category)

            # 保存DNS服务器
            for server in servers:
                value = (
                    f"{server.primary},{server.secondary}"
                    if server.secondary
                    else server.primary
                )
                self.config.set(category, server.name, value)

            logger.info(
                f"已保存 {len(servers)} 个DNS服务器到分类 {category}"
            )
            return True
        except Exception as e:
            logger.error(f"保存DNS服务器失败: {e}")
            return False

    def create_default_config(self) -> bool:
        """
        创建默认配置
        Returns:
            bool: 创建成功返回True，失败返回False
        """
        try:
            # 添加默认分类
            for category in ["公共IPv4", "公共IPv6"]:
                self.config.add_section(category)

            # 添加默认IPv4服务器
            for server in AppConfig.DEFAULT_IPV4[:3]:
                self.config.set( 
                    "公共IPv4",
                    server["name"],
                    f"{server['primary']},{server['secondary']}",
                )

            # 添加默认IPv6服务器
            for name, dns in AppConfig.DEFAULT_IPV6.items():
                self.config.set("公共IPv6", name, dns)

            # 保存配置
            return self.save_config()
        except Exception as e:
            logger.error(f"创建默认配置失败: {e}")
            return False

    def add_category(self, category: str) -> bool:
        """
        添加新分类

        Args:
            category: 分类名称

        Returns:
            bool: 添加成功返回True，失败返回False
        """
        if category == AppConfig.RESERVED_SECTION_NAME:
            logger.error(f"不能使用保留名称: {category}")
            return False

        if category in self.config.sections():
            logger.warning(f"分类已存在: {category}")
            return False

        try:
            self.config.add_section(category)
            self.save_config()
            logger.info(f"已添加分类: {category}")
            return True
        except Exception as e:
            logger.error(f"添加分类失败: {e}")
            return False

    def delete_category(self, category: str) -> bool:
        """
        删除分类

        Args:
            category: 分类名称

        Returns:
            bool: 删除成功返回True，失败返回False
        """
        if category == AppConfig.RESERVED_SECTION_NAME:
            logger.error(f"不能删除保留分类: {category}")
            return False

        if category not in self.config.sections():
            logger.warning(f"分类不存在: {category}")
            return False

        try:
            self.config.remove_section(category)
            self.save_config()
            logger.info(f"已删除分类: {category}")
            return True
        except Exception as e:
            logger.error(f"删除分类失败: {e}")
            return False

    def rename_category(self, old_name: str, new_name: str) -> bool:
        """
        重命名分类

        Args:
            old_name: 旧分类名称
            new_name: 新分类名称

        Returns:
            bool: 重命名成功返回True，失败返回False
        """
        if old_name == AppConfig.RESERVED_SECTION_NAME:
            logger.error(f"不能重命名保留分类: {old_name}")
            return False

        if new_name == AppConfig.RESERVED_SECTION_NAME:
            logger.error(f"不能使用保留名称: {new_name}")
            return False

        if old_name not in self.config.sections():
            logger.warning(f"分类不存在: {old_name}")
            return False

        if new_name in self.config.sections():
            logger.warning(f"分类已存在: {new_name}")
            return False

        try:
            # 复制旧分类的所有配置到新分类
            self.config.add_section(new_name)
            for key, value in self.config.items(old_name):
                self.config.set(new_name, key, value)

            # 删除旧分类
            self.config.remove_section(old_name)
            self.save_config()

            logger.info(f"已将分类 '{old_name}' 重命名为 '{new_name}'")
            return True
        except Exception as e:
            logger.error(f"重命名分类失败: {e}")
            return False