"""DNS配置处理器 - 处理DNS配置的业务逻辑"""

from typing import List, Dict, Optional

from config.config_manager import ConfigManager
from config.models import DNSServer, Category
from config.settings import AppConfig
from utils.logger import get_logger

logger = get_logger("DNSConfigHandler")


class DNSConfigHandler:
    """DNS配置处理器"""

    def __init__(self, config_manager: Optional[ConfigManager] = None):
        """
        初始化DNS配置处理器

        Args:
            config_manager: 配置管理器实例，如果为None则创建新实例
        """
        self.config_manager = config_manager or ConfigManager()
        logger.info("DNS配置处理器初始化完成")

    def load_categories(self) -> List[str]:
        """
        加载所有分类

        Returns:
            分类名称列表
        """
        try:
            self.config_manager.load_config()
            categories = self.config_manager.get_category_order()
            logger.info(f"加载了 {len(categories)} 个分类")
            return categories
        except Exception as e:
            logger.error(f"加载分类失败: {e}")
            return ["Ipv4_默认", "Ipv6_默认"]

    def load_category_servers(self, category: str) -> List[DNSServer]:
        """
        加载指定分类的DNS服务器

        Args:
            category: 分类名称

        Returns:
            DNS服务器列表
        """
        try:
            self.config_manager.load_config()
            servers = self.config_manager.get_category_dns_servers(category)
            logger.info(f"从分类 {category} 加载了 {len(servers)} 个DNS服务器")
            return servers
        except Exception as e:
            logger.error(f"加载分类DNS服务器失败: {e}")
            return []

    def save_category_servers(
        self, category: str, servers: List[DNSServer]
    ) -> bool:
        """
        保存分类的DNS服务器

        Args:
            category: 分类名称
            servers: DNS服务器列表

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            self.config_manager.load_config()
            success = self.config_manager.save_category_dns_servers(
                category, servers
            )
            if success:
                logger.info(f"成功保存 {len(servers)} 个DNS服务器到分类 {category}")
            return success
        except Exception as e:
            logger.error(f"保存分类DNS服务器失败: {e}")
            return False

    def add_category(self, category: str) -> bool:
        """
        添加新分类

        Args:
            category: 分类名称

        Returns:
            bool: 添加成功返回True，失败返回False
        """
        try:
            self.config_manager.load_config()
            success = self.config_manager.add_category(category)
            if success:
                logger.info(f"成功添加分类: {category}")
            return success
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
        try:
            self.config_manager.load_config()
            success = self.config_manager.delete_category(category)
            if success:
                logger.info(f"成功删除分类: {category}")
            return success
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
        try:
            self.config_manager.load_config()
            success = self.config_manager.rename_category(old_name, new_name)
            if success:
                logger.info(f"成功将分类 '{old_name}' 重命名为 '{new_name}'")
            return success
        except Exception as e:
            logger.error(f"重命名分类失败: {e}")
            return False

    def update_category_order(self, categories: List[str]) -> bool:
        """
        更新分类顺序

        Args:
            categories: 排序后的分类列表

        Returns:
            bool: 更新成功返回True，失败返回False
        """
        try:
            self.config_manager.load_config()
            success = self.config_manager.set_category_order(categories)
            if success:
                logger.info(f"成功更新分类顺序")
            return success
        except Exception as e:
            logger.error(f"更新分类顺序失败: {e}")
            return False

    def get_last_category(self) -> Optional[str]:
        """
        获取上次选择的分类

        Returns:
            上次选择的分类名称，如果不存在返回None
        """
        try:
            self.config_manager.load_config()
            return self.config_manager.get_last_category()
        except Exception as e:
            logger.error(f"获取上次选择的分类失败: {e}")
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
            self.config_manager.load_config()
            success = self.config_manager.set_last_category(category)
            if success:
                logger.info(f"成功设置上次选择的分类: {category}")
            return success
        except Exception as e:
            logger.error(f"设置上次选择的分类失败: {e}")
            return False

    def get_theme(self) -> str:
        """
        获取保存的主题

        Returns:
            主题名称
        """
        try:
            self.config_manager.load_config()
            return self.config_manager.get_theme()
        except Exception as e:
            logger.error(f"获取主题失败: {e}")
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
            self.config_manager.load_config()
            success = self.config_manager.set_theme(theme)
            if success:
                logger.info(f"成功设置主题: {theme}")
            return success
        except Exception as e:
            logger.error(f"设置主题失败: {e}")
            return False

    def create_default_config(self) -> bool:
        """
        创建默认配置

        Returns:
            bool: 创建成功返回True，失败返回False
        """
        try:
            success = self.config_manager.create_default_config()
            if success:
                logger.info("成功创建默认配置")
            return success
        except Exception as e:
            logger.error(f"创建默认配置失败: {e}")
            return False

    def move_servers_to_category(
        self, servers: List[DNSServer], from_category: str, to_category: str
    ) -> bool:
        """
        将DNS服务器从一个分类移动到另一个分类

        Args:
            servers: 要移动的DNS服务器列表
            from_category: 源分类名称
            to_category: 目标分类名称

        Returns:
            bool: 移动成功返回True，失败返回False
        """
        try:
            self.config_manager.load_config()

            # 获取源分类的服务器
            from_servers = self.config_manager.get_category_dns_servers(from_category)

            # 从源分类中移除指定的服务器
            from_servers = [
                s for s in from_servers if s.name not in {srv.name for srv in servers}
            ]

            # 保存源分类
            self.config_manager.save_category_dns_servers(from_category, from_servers)

            # 获取目标分类的服务器
            to_servers = self.config_manager.get_category_dns_servers(to_category)

            # 添加服务器到目标分类（避免重复）
            existing_names = {s.name for s in to_servers}
            for server in servers:
                if server.name not in existing_names:
                    to_servers.append(server)

            # 保存目标分类
            self.config_manager.save_category_dns_servers(to_category, to_servers)

            logger.info(
                f"成功将 {len(servers)} 个DNS服务器从 {from_category} 移动到 {to_category}"
            )
            return True

        except Exception as e:
            logger.error(f"移动DNS服务器失败: {e}")
            return False