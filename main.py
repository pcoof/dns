import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import PRIMARY, SECONDARY, SUCCESS, DANGER, WARNING, INFO
import dns.resolver
import time
import threading
import configparser
import platform
import ctypes
import pythoncom
import win32com.client
import os
import sys
import socket
from pathlib import Path
import webbrowser


# 应用程序常量
class AppConfig:
    WINDOW_TITLE = "DNS服务器测试工具"
    WINDOW_SIZE = (980, 500)
    CONFIG_FILE = "dns_servers.ini"
    FONT_FAMILY = "Microsoft YaHei"

    # 主题选项
    THEMES = [
        "Cyborg",
        "Darkly",
        "Superhero",
        "Vapor",
        "Solar",
        "Flatly",
        "Cosmo",
        "Journal",
        "Litera",
        "Lumen",
        "Minty",
        "Morph",
        "Pulse",
        "Sandstone",
        "Simplex",
        "Yeti",
        "United",
    ]
    # 深色主题列表
    DARK_THEMES = {"cyborg", "darkly", "superhero", "vapor", "solar"}

    # 默认DNS服务器
    DEFAULT_IPV4_SERVERS = [
        {
            "name": "US - Google Public DNS",
            "primary": "8.8.8.8",
            "secondary": "8.8.4.4",
        },
        {"name": "AU - Cloudflare", "primary": "1.1.1.1", "secondary": "1.0.0.1"},
        {
            "name": "US - OpenDNS",
            "primary": "208.67.222.222",
            "secondary": "208.67.220.220",
        },
        {"name": "CN - Aliyun", "primary": "223.5.5.5", "secondary": "223.6.6.6"},
        {
            "name": "CN - 114DNS",
            "primary": "114.114.114.114",
            "secondary": "114.114.115.115",
        },
        {
            "name": "CN - DNSPod",
            "primary": "119.29.29.29",
            "secondary": "182.254.116.116",
        },
    ]

    DEFAULT_IPV6_SERVERS = {
        "US - Google Public DNS": "2001:4860:4860::8888,2001:4860:4860::8844",
        "AU - Cloudflare": "2606:4700:4700::1111,2606:4700:4700::1001",
    }


class DNSTesterApp:
    def __init__(self, root):
        self.root = root
        self._init_window()
        self._init_data()
        self.create_widgets()
        self._load_initial_data()

    def _init_window(self):
        """初始化窗口设置"""
        self.root.title("DNS服务器测试工具")
        self.root.minsize(*AppConfig.WINDOW_SIZE)
        self.center_window(self.root, *AppConfig.WINDOW_SIZE)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _init_data(self):
        """初始化数据结构"""
        self.dns_servers = []
        self.test_results = {}
        self.network_connections = []
        self.current_ip = ""
        self.dns_categories = []
        self.current_category = "Ipv4_默认"

    def _load_initial_data(self):
        """加载初始数据"""
        self.load_default_config()
        self.load_theme_preference()
        self.load_network_connections()

    def create_widgets(self):
        """创建所有UI组件"""
        self._create_network_info_frame()
        self._create_control_frame()
        self._create_button_frame()
        self._create_status_frame()
        self._create_treeview_frame()

    def _create_network_info_frame(self):
        """创建网络信息显示区域"""
        network_frame = tb.Frame(self.root)
        network_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

        # 网络信息标签配置
        info_configs = [
            ("当前IP:", "ip_var", SUCCESS),
            ("默认网关:", "gateway_var", INFO),
            ("主DNS:", "primary_dns_var", PRIMARY),
            ("备用DNS:", "secondary_dns_var", WARNING),
        ]

        for label_text, var_name, style in info_configs:
            tb.Label(network_frame, text=label_text).pack(side=tk.LEFT, padx=(0, 5))
            var = tk.StringVar(value="获取中...")
            setattr(self, var_name, var)
            label = tb.Label(network_frame, textvariable=var, bootstyle=style)
            label.pack(side=tk.LEFT, padx=(0, 5))
            setattr(self, f"{var_name.replace('_var', '')}_label", label)

    def _create_control_frame(self):
        """创建控制区域"""
        category_frame = tb.Frame(self.root)
        category_frame.pack(fill=tk.X, padx=5, pady=5)
        # 网络设备选择
        self._create_network_device_controls(category_frame)
        # DNS类别选择
        self._create_category_controls(category_frame)
        # 主题选择
        self._create_theme_controls(category_frame)

    def _create_network_device_controls(self, parent):
        """创建网络设备控制组件"""
        tb.Label(parent, text="网络设备:", font=(AppConfig.FONT_FAMILY, 10)).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        self.network_var = tk.StringVar()
        self.network_combo = tb.Combobox(
            parent, textvariable=self.network_var, state="readonly", width=12
        )
        self.network_combo.pack(side=tk.LEFT, padx=(0, 10))
        tb.Button(
            parent, text="刷新", command=self.load_network_connections, bootstyle=INFO
        ).pack(side=tk.LEFT, padx=(0, 15))

    def _create_category_controls(self, parent):
        """创建DNS类别控制组件"""
        tb.Label(parent, text="DNS类别:", font=(AppConfig.FONT_FAMILY, 10)).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        self.category_var = tk.StringVar()
        self.category_combo = tb.Combobox(
            parent, textvariable=self.category_var, state="readonly", width=15
        )
        self.category_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.category_combo.bind("<<ComboboxSelected>>", self.on_category_changed)
        tb.Button(
            parent, text="管理分类", command=self.manage_categories, bootstyle=SECONDARY
        ).pack(side=tk.LEFT, padx=5)

    def _create_theme_controls(self, parent):
        """创建主题控制组件"""
        tb.Label(parent, text="主题:", font=(AppConfig.FONT_FAMILY, 10)).pack(
            side=tk.LEFT, padx=(15, 5)
        )
        self.theme_var = tk.StringVar()
        self.theme_combo = tb.Combobox(
            parent, textvariable=self.theme_var, state="readonly", width=12
        )
        self.theme_combo["values"] = AppConfig.THEMES
        self.theme_combo.set("Darkly")
        self.theme_combo.pack(side=tk.LEFT, padx=(0, 5))
        self.theme_combo.bind("<<ComboboxSelected>>", self.on_theme_changed)

    def _create_button_frame(self):
        """创建按钮区域"""
        button_frame = tb.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        buttons = [
            ("添加DNS", self.add_dns, SUCCESS),
            ("删除选中", self.remove_dns, DANGER),
            ("开始测试", self.start_test, PRIMARY),
            ("清理DNS缓存", self.clear_dns_cache, INFO),
            ("默认DNS", self.reset_to_dhcp, WARNING),
        ]
        # 配置grid权重，使按钮平均分布
        for i in range(len(buttons)):
            button_frame.grid_columnconfigure(i, weight=1)
        # 创建按钮
        for i, (text, command, style) in enumerate(buttons):
            tb.Button(button_frame, text=text, command=command, bootstyle=style).grid(
                row=0, column=i, padx=2, sticky="ew"
            )

    def _create_status_frame(self):
        """创建状态栏"""
        status_frame = tb.Frame(self.root, bootstyle=SECONDARY)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        # 左侧状态信息
        self.status_var = tk.StringVar(value="就绪 - 管理员权限")
        self.status_bar = tb.Label(
            status_frame, textvariable=self.status_var, bootstyle=SECONDARY
        )
        self.status_bar.pack(side=tk.LEFT, padx=10, pady=5)
        # 右侧GitHub链接
        self._create_github_link(status_frame)

    def _create_github_link(self, parent):
        """创建GitHub链接"""
        github_frame = tb.Frame(parent)
        github_frame.pack(side=tk.RIGHT, padx=10, pady=5)
        github_label = tb.Label(
            github_frame, text="⭐ GitHub", bootstyle="INFO", cursor="hand2"
        )
        github_label.pack(side=tk.RIGHT)
        github_label.bind("<Button-1>", self.open_github)

    def _create_treeview_frame(self):
        """创建列表视图"""
        tree_frame = tb.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))
        # 创建Treeview
        columns = ("name", "primary", "secondary", "latency", "status")
        self.tree = tb.Treeview(
            tree_frame, columns=columns, show="headings", bootstyle=INFO, height=30
        )
        self._setup_treeview_style()
        self._setup_treeview_columns()
        self._setup_treeview_scrollbar(tree_frame)
        self._setup_treeview_events()

    def _setup_treeview_style(self):
        """设置Treeview样式"""
        style = tb.Style()
        style.configure(
            "Treeview",
            relief="solid",
            borderwidth=1,
            rowheight=30,
        )
        style.configure("Treeview.Heading", relief="solid", borderwidth=1)
        self.tree.configure(style="Treeview")

    def _setup_treeview_columns(self):
        """设置Treeview列"""
        column_configs = [
            ("name", "名称", 100, tk.W),
            ("primary", "主DNS", 150, tk.W),
            ("secondary", "备用DNS", 150, tk.W),
            ("latency", "延迟(ms)", 80, tk.CENTER),
            ("status", "状态", 60, tk.CENTER),
        ]
        for col_id, heading, width, anchor in column_configs:
            self.tree.heading(col_id, text=heading)
            self.tree.column(
                col_id,
                width=width,
                anchor=anchor,
                minwidth=width if anchor == tk.CENTER else 100,
            )

    def _setup_treeview_scrollbar(self, parent):
        """设置Treeview滚动条"""
        scrollbar = tb.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _setup_treeview_events(self):
        """设置Treeview事件"""
        self.popup_menu = tk.Menu(self.root, tearoff=0)
        self.tree.bind("<Button-3>", self.show_popup)
        self.tree.bind("<Double-1>", self.on_double_click)

    def open_github(self, event):
        """打开GitHub链接"""
        try:
            webbrowser.open("https://github.com/pcoof/dns-tester")
            self.show_notification("已在浏览器中打开GitHub页面", INFO)
        except Exception as e:
            print(f"打开GitHub链接失败: {e}")
            self.show_notification("打开GitHub链接失败", WARNING)

    def _set_title_bar_theme(self, is_dark_theme):
        """设置Windows标题栏主题"""
        try:
            import ctypes

            # 确保窗口已经完全创建
            self.root.update_idletasks()
            try:
                import sys

                # 使用GetParent获取正确的窗口句柄
                hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                # 如果GetParent返回0，使用原始窗口句柄
                if hwnd == 0:
                    hwnd = self.root.winfo_id()
            except Exception:
                return
            try:
                # 设置主题值
                value = 1 if is_dark_theme else 0
                if sys.getwindowsversion().build >= 22000:  # Windows 11
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                else:  # Windows 10
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 19
                # 使用theme.py中的方法：c_int类型
                result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(ctypes.c_int(value)),
                    ctypes.sizeof(ctypes.c_int),
                )
                if result == 0:  # S_OK
                    # 强制窗口重绘
                    self.root.update()
                    success = True
                else:
                    success = False
            except Exception:
                pass
        except Exception as e:
            print(f"❌ 设置标题栏主题失败: {e}")

    def _is_dark_theme(self, theme_name):
        """判断主题是否为深色主题"""
        return theme_name.lower() in AppConfig.DARK_THEMES

    @staticmethod
    def _is_ip(ip, version=None):
        """
        检查是否为有效的IP地址
        参数:
            ip: 待检查的IP地址字符串
            version: 可选，指定IP版本(4或6)，默认检查是否为任意有效IP
        返回:
            bool: 符合指定版本则返回True，否则返回False
        """

        # 内部IPv4检查逻辑
        def check_ipv4():
            if not ip:
                return True  # 空字符串视为有效IPv4(保持原逻辑)
            try:
                parts = ip.split(".")
                return len(parts) == 4 and all(0 <= int(part) <= 255 for part in parts)
            except (ValueError, AttributeError):
                return False

        # 内部IPv6检查逻辑
        def check_ipv6():
            if not ip:
                return False  # 空字符串视为无效IPv6(保持原逻辑)
            try:
                socket.inet_pton(socket.AF_INET6, ip)
                return True
            except (socket.error, AttributeError):
                return False

        # 根据版本参数返回对应检查结果
        if version == 4:
            return check_ipv4()
        elif version == 6:
            return check_ipv6()
        elif version is None:
            # 检查是否为任意有效IP(空字符串返回False)
            if not ip:
                return False
            return check_ipv4() or check_ipv6()
        else:
            # 无效版本参数返回False
            return False

    def _get_network_adapters_info(self):
        """获取网络适配器信息"""
        try:
            pythoncom.CoInitialize()
            wmi = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")
            if not wmi:
                return []
            adapters = wmi.ExecQuery(
                "SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True"
            )
            adapter_info_map = {}
            for adapter in adapters:
                adapter_info_map[str(adapter.DeviceID)] = {
                    "name": adapter.Name,
                    "net_connection_id": adapter.NetConnectionID,  # 这是友好名称，如"以太网 2"
                }
            # 查询启用的网络适配器配置
            configs = wmi.ExecQuery(
                "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
            )
            network_info = []
            for config in configs:
                adapter_info = adapter_info_map.get(str(config.Index), {})
                # 优先使用NetConnectionID（友好名称），如果没有则使用Name
                adapter_name = (
                    adapter_info.get("net_connection_id")
                    or adapter_info.get("name")
                    or "未识别适配器"
                )
                # 分离IPv4和IPv6地址
                ip_addresses = list(config.IPAddress) if config.IPAddress else []
                ipv4_addresses = [ip for ip in ip_addresses if not self._is_ip(ip, 6)]
                ipv6_addresses = [ip for ip in ip_addresses if self._is_ip(ip, 6)]
                # 分离IPv4和IPv6网关
                gateways = (
                    list(config.DefaultIPGateway) if config.DefaultIPGateway else []
                )
                ipv4_gateways = [gw for gw in gateways if not self._is_ip(gw, 6)]
                ipv6_gateways = [gw for gw in gateways if self._is_ip(gw, 6)]
                info = {
                    "name": adapter_name,
                    "index": config.Index,
                    "mac": config.MACAddress,
                    "ipv4_addresses": ipv4_addresses,
                    "ipv6_addresses": ipv6_addresses,
                    "ipv4_gateways": ipv4_gateways,
                    "ipv6_gateways": ipv6_gateways,
                    "dns_servers": list(config.DNSServerSearchOrder)
                    if config.DNSServerSearchOrder
                    else [],
                    "dhcp_enabled": config.DHCPEnabled,
                }
                network_info.append(info)
            return network_info
        except Exception:
            return []
        finally:
            pythoncom.CoUninitialize()

    def center_window(self, window, width, height):
        """将窗口居中显示在屏幕上"""
        # 获取屏幕宽度和高度
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        # 计算窗口的坐标
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        # 设置窗口大小和位置
        window.geometry(f"{width}x{height}+{x}+{y}")

    def on_closing(self):
        """窗口关闭时的处理"""
        try:
            print("程序正在关闭，保存配置...")
            # 保存当前配置
            self.auto_save_current_category()
            # 保存当前选择的类别
            self.auto_save_selection()
            print("配置保存完成，程序退出")
        except Exception as e:
            print(f"关闭时保存配置失败: {e}")
        finally:
            # 销毁窗口
            self.root.destroy()

    def show_notification(self, message, style=INFO):
        """显示应用内通知到状态栏，3秒后自动恢复"""
        if not hasattr(self, "_original_status"):
            self._original_status = self.status_var.get()

        prefix, color = self._get_notification_style(style)
        self.status_var.set(f"{prefix}{message}")
        self.status_bar.configure(foreground=color)
        self.root.after("3000", self.hide_notification)

    def _get_notification_style(self, style):
        """获取通知样式"""
        style_map = {
            SUCCESS: ("✓ ", "green"),
            DANGER: ("✗ ", "red"),
            WARNING: ("⚠ ", "orange"),
            INFO: ("ℹ ", "blue"),
        }
        return style_map.get(style, ("ℹ ", "blue"))

    def hide_notification(self):
        """恢复状态栏原始显示"""
        if hasattr(self, "_original_status"):
            self.status_var.set(self._original_status)
            delattr(self, "_original_status")
        else:
            self.status_var.set("就绪 - 管理员权限")
        self.status_bar.configure(foreground="")

    def update_status(self, message):
        """更新状态栏（非通知）"""
        # 如果当前没有显示通知，直接更新
        if not hasattr(self, "_original_status"):
            self.status_var.set(message)
        else:
            # 如果正在显示通知，更新保存的原始状态
            self._original_status = message

    def get_local_ip(self):
        """获取本机局域网IP地址"""
        try:
            # 使用WMI获取网络适配器信息
            adapters = self._get_network_adapters_info()
            for adapter in adapters:
                if adapter["ipv4_addresses"]:
                    for ip in adapter["ipv4_addresses"]:
                        # 排除回环地址和APIPA地址
                        if not ip.startswith(("127.", "169.254.")):
                            return ip

            # 备选方案：使用socket方法
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                s.connect(("8.8.8.8", 80))
                local_ip = s.getsockname()[0]
                return local_ip
        except Exception:
            try:
                # 最后备选方案：获取主机名对应的IP
                hostname = socket.gethostname()
                local_ip = socket.gethostbyname(hostname)
                return local_ip
            except Exception:
                return "无法获取"

    def get_default_gateway(self):
        """获取默认网关地址"""
        try:
            # 使用WMI获取网络适配器信息
            adapters = self._get_network_adapters_info()
            for adapter in adapters:
                if adapter["ipv4_gateways"]:
                    # 返回第一个IPv4网关
                    return adapter["ipv4_gateways"][0]
        except Exception as e:
            print(f"获取默认网关失败: {e}")
        return "无法获取"

    def get_current_dns_servers(self):
        """获取当前系统DNS服务器"""
        try:
            # 使用WMI获取网络适配器信息
            adapters = self._get_network_adapters_info()
            for adapter in adapters:
                if adapter["dns_servers"]:
                    # 返回前两个DNS服务器
                    dns_list = adapter["dns_servers"][:2]
                    # 确保返回两个元素
                    if len(dns_list) == 1:
                        dns_list.append("未设置")
                    return dns_list
        except Exception as e:
            print(f"获取当前DNS服务器失败: {e}")
        return ["无法获取", ""]

    def load_default_config(self):
        """加载默认配置"""
        config_path = Path(AppConfig.CONFIG_FILE)
        if config_path.exists():
            self.load_config()
        else:
            self._create_default_config()

    def _create_default_config(self):
        """创建默认配置"""
        self.dns_categories = ["Ipv4_默认", "Ipv6_默认"]
        self.current_category = "Ipv4_默认"
        self.category_combo["values"] = self.dns_categories
        self.category_combo.set(self.current_category)

        self.dns_servers = AppConfig.DEFAULT_IPV4_SERVERS.copy()
        self.update_treeview()
        self.show_notification("已加载默认DNS配置", INFO)

    def load_config(self):
        """加载配置文件"""
        config = self._get_config_parser()
        try:
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            categories = self._get_categories_from_config(config)

            if not categories:
                self._create_default_categories(config)
                categories = ["Ipv4_默认", "Ipv6_默认"]

            self._load_category_order(config, categories)
            self._load_last_selection()
            self._update_category_combo()
            self.load_category_dns(self.current_category)

        except Exception as e:
            self.show_notification(f"自动加载配置失败: {str(e)}", DANGER)

    def _get_config_parser(self):
        """获取配置解析器"""
        config = configparser.ConfigParser()
        config.optionxform = str  # 区分大小写
        return config

    def _get_categories_from_config(self, config):
        """从配置中获取类别列表"""
        return [section for section in config.sections() if section != "Main"]

    def _create_default_categories(self, config):
        """创建默认类别"""
        for category in ["Ipv4_默认", "Ipv6_默认"]:
            config.add_section(category)

        # 添加默认IPv4服务器
        for server in AppConfig.DEFAULT_IPV4_SERVERS[:3]:  # 只添加前3个
            config.set(
                "Ipv4_默认",
                server["name"],
                f"{server['primary']},{server['secondary']}",
            )

        # 添加默认IPv6服务器
        for name, dns in AppConfig.DEFAULT_IPV6_SERVERS.items():
            config.set("Ipv6_默认", name, dns)

        self._save_config_file(config)
        self.show_notification("已创建默认DNS分类", SUCCESS)

    def _load_category_order(self, config, categories):
        """加载类别顺序"""
        if "Main" in config.sections() and config.has_option("Main", "category_order"):
            saved_order = config.get("Main", "category_order").split(",")
            self.dns_categories = [cat for cat in saved_order if cat in categories]
            # 添加新类别
            for cat in categories:
                if cat not in self.dns_categories:
                    self.dns_categories.append(cat)
        else:
            self.dns_categories = categories

    def _load_last_selection(self):
        """加载上次选择的类别"""
        self.load_last_selection()
        if self.current_category not in self.dns_categories and self.dns_categories:
            self.current_category = self.dns_categories[0]

    def _update_category_combo(self):
        """更新类别下拉框"""
        if self.dns_categories:
            self.category_combo["values"] = self.dns_categories
            self.category_combo.set(self.current_category)

    def _save_config_file(self, config):
        """保存配置文件"""
        with open(AppConfig.CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)

    def load_category_dns(self, category):
        """加载指定类别的DNS服务器"""
        config = self._get_config_parser()
        try:
            config.read("dns_servers.ini", encoding="utf-8")
            self.dns_servers = []
            # 如果类别不存在，创建该类别
            if category not in config.sections():
                config.add_section(category)
                # 如果是默认类别，添加一些默认DNS服务器
                if category == "Ipv4_默认":
                    default_servers = {
                        "US - Google Public DNS": "8.8.8.8,8.8.4.4",
                        "AU - Cloudflare": "1.1.1.1,1.0.0.1",
                        "CN - Aliyun": "223.5.5.5,223.6.6.6",
                    }
                    for name, dns in default_servers.items():
                        config.set(category, name, dns)
                elif category == "Ipv6_默认":
                    default_servers = {
                        "US - Google Public DNS": "2001:4860:4860::8888,2001:4860:4860::8844",
                        "AU - Cloudflare": "2606:4700:4700::1111,2606:4700:4700::1001",
                    }
                    for name, dns in default_servers.items():
                        config.set(category, name, dns)
                # 保存配置文件
                self._save_config_file(config)
                self.show_notification(f"已创建类别: {category}", SUCCESS)
            # 判断是IPv4还是IPv6类别
            is_ipv6_category = category.startswith("Ipv6_")
            filtered_addresses = []
            # 解析新格式的DNS配置
            for key, value in config.items(category):
                if not value.strip():
                    continue
                # 解析格式: name=primary,secondary[,other] 或 name=primary
                dns_parts = value.split(",")
                if len(dns_parts) >= 2:
                    primary = dns_parts[0].strip()
                    secondary = dns_parts[1].strip()
                    # 忽略第三个及以后的参数（如True/False标志）
                elif len(dns_parts) == 1:
                    primary = dns_parts[0].strip()
                    secondary = ""
                else:
                    continue
                # 过滤掉非IP地址的参数（如"True", "False"等）
                if primary and not DNSTesterApp._is_ip(primary):
                    continue
                # 验证备用DNS地址
                if secondary and not DNSTesterApp._is_ip(secondary):
                    secondary = ""

                # 根据类别过滤地址
                if is_ipv6_category:
                    # IPv6类别：过滤掉IPv4地址
                    if (
                        primary
                        and self._is_ip(primary, 4)
                        and not self._is_ip(primary, 6)
                    ):
                        primary = ""
                        filtered_addresses.append(f"{key} 主DNS(IPv4)")
                    if (
                        secondary
                        and self._is_ip(secondary, 4)
                        and not self._is_ip(secondary, 6)
                    ):
                        secondary = ""
                        filtered_addresses.append(f"{key} 备用DNS(IPv4)")
                else:
                    # IPv4类别：过滤掉IPv6地址
                    if primary and not self._is_ip(primary, 4):
                        primary = ""
                        filtered_addresses.append(f"{key} 主DNS(IPv6)")
                    if secondary and not self._is_ip(secondary, 4):
                        secondary = ""
                        filtered_addresses.append(f"{key} 备用DNS(IPv6)")
                if primary:  # 只添加有主DNS的服务器
                    self.dns_servers.append(
                        {
                            "name": key,
                            "primary": primary,
                            "secondary": secondary,
                        }
                    )
            self.update_treeview()
            # 显示加载结果通知
            if filtered_addresses:
                addr_type = "IPv4" if is_ipv6_category else "IPv6"
                filtered_msg = f"已加载 {category}，过滤了 {len(filtered_addresses)} 个{addr_type}地址"
                self.show_notification(filtered_msg, WARNING)
            else:
                self.show_notification(
                    f"已加载 {category} ({len(self.dns_servers)} 个DNS服务器)", SUCCESS
                )
        except Exception as e:
            self.show_notification(f"加载类别失败: {str(e)}", DANGER)

    def on_category_changed(self, event=None):
        """当DNS类别选择改变时的回调"""
        selected_category = self.category_var.get()
        if selected_category and selected_category != self.current_category:
            self.current_category = selected_category
            self.test_results = {}  # 清空测试结果
            self.load_category_dns(selected_category)
            # 自动保存当前选择
            self.auto_save_selection()

    def on_theme_changed(self, event=None):
        """当主题选择改变时的回调"""
        selected_theme = self.theme_var.get()
        if selected_theme:
            try:
                # 获取当前主题名称（小写）
                theme_name = selected_theme.lower()
                # 应用新主题
                style = tb.Style()
                style.theme_use(theme_name)
                # 重新应用自定义的Treeview样式
                self._reapply_treeview_styles()
                # 根据主题设置标题栏颜色（延迟执行确保主题已应用）
                is_dark = self._is_dark_theme(theme_name)
                self.root.after(50, lambda: self._set_title_bar_theme(is_dark))
                # 显示成功通知
                theme_type = "深色" if is_dark else "浅色"
                self.show_notification(
                    f"已切换到 {selected_theme} 主题 ({theme_type})", SUCCESS
                )
                # 保存主题选择到配置文件
                self.save_theme_preference(theme_name)
            except Exception:
                self.show_notification(f"切换到 {selected_theme} 主题失败", DANGER)
                # 恢复到之前的主题
                self.theme_combo.set("Darkly")

    def _reapply_treeview_styles(self):
        """重新应用Treeview的自定义样式"""
        try:
            style = tb.Style()
            # 重新配置Treeview基本样式，保持大行高
            style.configure(
                "Treeview",
                relief="solid",
                borderwidth=1,
                rowheight=30,  # 保持30像素的行高
            )
            # # 重新配置表头样式
            style.configure(
                "Treeview.Heading",
                relief="solid",
                borderwidth=1,
                font=(AppConfig.FONT_FAMILY, 9, "bold"),
            )
            print("已重新应用Treeview自定义样式")
        except Exception as e:
            print(f"重新应用Treeview样式失败: {e}")

    def refresh_categories(self):
        """刷新DNS类别列表"""
        self.load_config()

    def manage_categories(self):
        """管理DNS分类"""
        manage_dialog = tk.Toplevel(self.root)
        manage_dialog.title("管理DNS分类")
        manage_dialog.geometry("500x500")
        manage_dialog.resizable(True, True)
        manage_dialog.transient(self.root)
        manage_dialog.grab_set()
        # 居中显示
        manage_dialog.update_idletasks()
        width = manage_dialog.winfo_width()
        height = manage_dialog.winfo_height()
        x = (self.root.winfo_width() // 2) - (width // 2) + self.root.winfo_x()
        y = (self.root.winfo_height() // 2) - (height // 2) + self.root.winfo_y()
        manage_dialog.geometry(f"+{x}+{y}")
        # 创建主框架
        main_frame = tb.Frame(manage_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        # 标题
        tb.Label(
            main_frame, text="DNS分类管理", font=(AppConfig.FONT_FAMILY, 12, "bold")
        ).pack(pady=(0, 10))
        # 分类列表和排序按钮容器
        list_container = tb.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        # 分类列表框架
        list_frame = tb.Frame(list_container)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 创建列表框
        category_listbox = tk.Listbox(list_frame, font=(AppConfig.FONT_FAMILY, 10))
        category_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 滚动条
        scrollbar = tb.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=category_listbox.yview
        )
        category_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # 排序按钮框架
        sort_frame = tb.Frame(list_container)
        sort_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))

        # 填充分类列表
        def refresh_category_list():
            category_listbox.delete(0, tk.END)
            for category in self.dns_categories:
                category_listbox.insert(tk.END, category)

        # 上移分类
        def move_up():
            selection = category_listbox.curselection()
            if not selection or selection[0] == 0:
                return
            idx = selection[0]
            category = self.dns_categories[idx]
            self.dns_categories.pop(idx)
            self.dns_categories.insert(idx - 1, category)
            refresh_category_list()
            category_listbox.selection_set(idx - 1)
            category_listbox.see(idx - 1)
            # 更新下拉框
            self.category_combo["values"] = self.dns_categories
            # 保存分类顺序
            self.save_category_order()

        # 下移分类
        def move_down():
            selection = category_listbox.curselection()
            if not selection or selection[0] == len(self.dns_categories) - 1:
                return
            idx = selection[0]
            category = self.dns_categories[idx]
            self.dns_categories.pop(idx)
            self.dns_categories.insert(idx + 1, category)
            refresh_category_list()
            category_listbox.selection_set(idx + 1)
            category_listbox.see(idx + 1)
            # 更新下拉框
            self.category_combo["values"] = self.dns_categories
            # 保存分类顺序
            self.save_category_order()

        # 添加排序按钮
        tb.Button(
            sort_frame, text="↑", width=3, command=move_up, bootstyle=PRIMARY
        ).pack(pady=(0, 5))
        tb.Button(
            sort_frame, text="↓", width=3, command=move_down, bootstyle=PRIMARY
        ).pack()
        refresh_category_list()
        # 按钮框架
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def add_category():
            new_category = tk.simpledialog.askstring("添加分类", "请输入新分类名称:")
            if new_category and new_category.strip():
                new_category = new_category.strip()
                # 确保分类名称格式正确
                if new_category == "Main":
                    messagebox.showwarning("警告", "不能使用保留名称 'Main'")
                    return
                if new_category not in self.dns_categories:
                    # 添加到配置文件
                    config = self._get_config_parser()
                    config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
                    if new_category not in config.sections():
                        config.add_section(new_category)
                        self._save_config_file(config)
                        self.dns_categories.append(new_category)
                        refresh_category_list()
                        self.category_combo["values"] = self.dns_categories
                        self.show_notification(f"已添加分类: {new_category}", SUCCESS)
                        # 保存分类顺序
                        self.save_category_order()
                else:
                    messagebox.showwarning("警告", "分类已存在")

        def delete_category():
            selection = category_listbox.curselection()
            if not selection:
                messagebox.showinfo("提示", "请先选择要删除的分类")
                return
            category_to_delete = category_listbox.get(selection[0])
            if category_to_delete == self.current_category:
                messagebox.showwarning("警告", "不能删除当前正在使用的分类")
                return
            if messagebox.askyesno(
                "确认",
                f"确定要删除分类 '{category_to_delete}' 吗？\n这将删除该分类下的所有DNS服务器。",
            ):
                # 从配置文件删除
                config = self._get_config_parser()
                config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
                if category_to_delete in config.sections():
                    config.remove_section(category_to_delete)
                    self._save_config_file(config)
                    self.dns_categories.remove(category_to_delete)
                    refresh_category_list()
                    self.category_combo["values"] = self.dns_categories
                    self.show_notification(f"已删除分类: {category_to_delete}", WARNING)

        def rename_category():
            selection = category_listbox.curselection()
            if not selection:
                messagebox.showinfo("提示", "请先选择要重命名的分类")
                return
            old_category = category_listbox.get(selection[0])
            if old_category == self.current_category:
                messagebox.showwarning("警告", "不能重命名当前正在使用的分类")
                return
            new_name = tk.simpledialog.askstring(
                "重命名分类", "请输入新的分类名称:", initialvalue=old_category
            )
            if new_name and new_name.strip() and new_name.strip() != old_category:
                new_name = new_name.strip()
                if new_name not in self.dns_categories:
                    # 重命名配置文件中的分类
                    config = self._get_config_parser()
                    config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
                    if old_category in config.sections():
                        # 复制旧分类的所有配置到新分类
                        config.add_section(new_name)
                        for key, value in config.items(old_category):
                            config.set(new_name, key, value)
                        # 删除旧分类
                        config.remove_section(old_category)
                        self._save_config_file(config)
                        # 更新内存中的分类列表
                        index = self.dns_categories.index(old_category)
                        self.dns_categories[index] = new_name
                        refresh_category_list()
                        self.category_combo["values"] = self.dns_categories
                        self.show_notification(
                            f"已将分类 '{old_category}' 重命名为 '{new_name}'", SUCCESS
                        )
                else:
                    messagebox.showwarning("警告", "新分类名称已存在")

        tb.Button(
            button_frame, text="添加分类", command=add_category, bootstyle=SUCCESS
        ).pack(side=tk.LEFT, padx=5)
        tb.Button(
            button_frame, text="重命名分类", command=rename_category, bootstyle=INFO
        ).pack(side=tk.LEFT, padx=5)
        tb.Button(
            button_frame, text="删除分类", command=delete_category, bootstyle=DANGER
        ).pack(side=tk.LEFT, padx=5)
        tb.Button(
            button_frame,
            text="关闭",
            command=manage_dialog.destroy,
            bootstyle=SECONDARY,
        ).pack(side=tk.RIGHT, padx=5)

    def auto_save_selection(self):
        """自动保存当前类别选择"""
        try:
            config = self._get_config_parser()
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            # 确保Main节存在
            if "Main" not in config.sections():
                config.add_section("Main")
            # 保存当前选择的类别
            config.set("Main", "last_category", self.current_category)
            self._save_config_file(config)
        except Exception as e:
            print(f"自动保存选择失败: {e}")

    def save_theme_preference(self, theme_name):
        """保存主题偏好到配置文件"""
        try:
            config = self._get_config_parser()
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            if "Main" not in config.sections():
                config.add_section("Main")
            config.set("Main", "theme", theme_name)
            self._save_config_file(config)
        except Exception as e:
            print(f"保存主题偏好失败: {e}")

    def load_theme_preference(self):
        """从配置文件加载主题偏好"""
        try:
            config = self._get_config_parser()
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            if "Main" in config.sections() and config.has_option("Main", "theme"):
                theme_name = config.get("Main", "theme")
                # 将主题名称首字母大写
                theme_display = theme_name.capitalize()
                self.theme_combo.set(theme_display)
                # 延迟设置标题栏主题，确保窗口完全加载
                is_dark = self._is_dark_theme(theme_name)
                self.root.after(100, lambda: self._set_title_bar_theme(is_dark))
                return theme_name
            else:
                # 默认主题也设置标题栏
                self.root.after(
                    100, lambda: self._set_title_bar_theme(True)
                )  # darkly是深色主题
                return "darkly"  # 默认主题
        except Exception:
            # 默认主题也设置标题栏
            self.root.after(100, lambda: self._set_title_bar_theme(True))
            return "darkly"

    def save_category_order(self):
        """保存分类顺序"""
        try:
            config = self._get_config_parser()
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            # 确保Main节存在
            if "Main" not in config.sections():
                config.add_section("Main")
            # 保存分类顺序
            config.set("Main", "category_order", ",".join(self.dns_categories))
            self._save_config_file(config)
        except Exception:
            pass

    def load_last_selection(self):
        """加载上次选择的类别"""
        try:
            config = self._get_config_parser()
            config.read(AppConfig.CONFIG_FILE, encoding="utf-8")
            if "Main" in config.sections():
                # 获取上次选择的类别，如果不存在则使用None
                last_category = config.get("Main", "last_category", fallback=None)
                if last_category and last_category in self.dns_categories:
                    self.current_category = last_category
                    return
            # 如果没有保存的类别或保存的类别不存在，选择第一个可用的类别
            if self.dns_categories:
                self.current_category = self.dns_categories[0]
            else:
                self.current_category = "Ipv4_默认"  # 作为最后的备选
        except Exception as e:
            print(f"加载上次选择失败: {e}")
            # 出错时选择第一个可用的类别
            if self.dns_categories:
                self.current_category = self.dns_categories[0]
            else:
                self.current_category = "Ipv4_默认"

    def auto_save_current_category(self):
        """自动保存当前类别的DNS配置"""
        try:
            config = self._get_config_parser()
            config_path = os.path.abspath(AppConfig.CONFIG_FILE)
            print(f"正在保存配置到: {config_path}")
            # 读取现有配置
            self._load_existing_config(config, config_path)
            # 更新当前类别
            self._update_category_in_config(config)
            # 保存DNS服务器
            saved_count = self._save_dns_servers_to_config(config)
            # 写入文件
            self._write_config_to_file(config, config_path)
            print(f"配置保存成功: {saved_count} 个DNS服务器已保存")
            self.show_notification(f"已保存 {saved_count} 个DNS服务器", SUCCESS)

        except PermissionError as e:
            self._handle_permission_error(e)
        except Exception as e:
            self._handle_save_error(e)

    def _load_existing_config(self, config, config_path):
        """加载现有配置"""
        if os.path.exists(config_path):
            config.read(config_path, encoding="utf-8")
            print(f"已读取现有配置，包含 {len(config.sections())} 个分类")
        else:
            print("配置文件不存在，将创建新文件")

    def _update_category_in_config(self, config):
        """更新配置中的类别"""
        if self.current_category not in config.sections():
            config.add_section(self.current_category)
            print(f"创建新分类: {self.current_category}")
        else:
            config.remove_section(self.current_category)
            config.add_section(self.current_category)
            print(f"清空并重建分类: {self.current_category}")

    def _save_dns_servers_to_config(self, config):
        """保存DNS服务器到配置"""
        saved_count = 0
        for server in self.dns_servers:
            value = (
                f"{server['primary']},{server['secondary']}"
                if server["secondary"]
                else server["primary"]
            )
            config.set(self.current_category, server["name"], value)
            saved_count += 1

        print(f"准备保存 {saved_count} 个DNS服务器到分类 {self.current_category}")
        return saved_count

    def _write_config_to_file(self, config, config_path):
        """写入配置到文件"""
        with open(config_path, "w", encoding="utf-8") as f:
            config.write(f)

    def _handle_permission_error(self, e):
        """处理权限错误"""
        error_msg = f"配置保存失败：没有写入权限 - {e}"
        print(error_msg)
        self.show_notification("保存失败：权限不足", DANGER)

    def _handle_save_error(self, e):
        """处理保存错误"""
        error_msg = f"自动保存配置失败: {e}"
        print(error_msg)
        self.show_notification(f"保存失败: {str(e)}", DANGER)

    def save_config(self):
        """保存当前类别的DNS配置"""
        config = self._get_config_parser()
        try:
            # 获取配置文件的完整路径
            config_path = os.path.abspath("dns_servers.ini")
            print(f"手动保存配置到: {config_path}")

            # 先读取现有配置
            if os.path.exists(config_path):
                config.read(config_path, encoding="utf-8")

            # 确保当前类别存在
            if self.current_category not in config.sections():
                config.add_section(self.current_category)
            else:
                # 清空当前类别的配置
                config.remove_section(self.current_category)
                config.add_section(self.current_category)

            # 保存当前DNS服务器到当前类别
            saved_count = 0
            for server in self.dns_servers:
                if server["secondary"]:
                    value = f"{server['primary']},{server['secondary']}"
                else:
                    value = server["primary"]
                config.set(self.current_category, server["name"], value)
                saved_count += 1

            # 写入配置文件
            with open(config_path, "w", encoding="utf-8") as f:
                config.write(f)

            print(f"手动保存成功: {saved_count} 个DNS服务器")
            self.show_notification(
                f"已保存 {self.current_category} 配置 ({saved_count} 个DNS)", SUCCESS
            )

        except PermissionError as e:
            error_msg = f"保存失败：没有写入权限 - {e}"
            print(error_msg)
            self.show_notification("保存失败：权限不足", DANGER)
        except Exception as e:
            error_msg = f"保存配置失败: {e}"
            print(error_msg)
            self.show_notification(f"保存配置失败: {str(e)}", DANGER)

    def add_dns(self):
        dns_dialog = tk.Toplevel(self.root)
        dns_dialog.title("添加DNS服务器")
        dns_dialog.geometry("400x260")
        dns_dialog.resizable(False, False)
        dns_dialog.transient(self.root)
        dns_dialog.grab_set()
        # 居中显示
        dns_dialog.update_idletasks()
        width = dns_dialog.winfo_width()
        height = dns_dialog.winfo_height()
        x = (self.root.winfo_width() // 2) - (width // 2) + self.root.winfo_x()
        y = (self.root.winfo_height() // 2) - (height // 2) + self.root.winfo_y()
        dns_dialog.geometry(f"+{x}+{y}")
        # 创建主框架
        main_frame = tb.Frame(dns_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        # 创建输入框
        tb.Label(main_frame, text="分类:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=0, column=0, padx=5, pady=10, sticky=tk.W
        )
        category_var = tk.StringVar()
        category_combo = tb.Combobox(
            main_frame,
            textvariable=category_var,
            width=22,
            font=(AppConfig.FONT_FAMILY, 9),
            state="readonly",
        )
        category_combo["values"] = self.dns_categories
        category_combo.set(self.current_category)
        category_combo.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W + tk.E)
        tb.Label(main_frame, text="名称:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=1, column=0, padx=5, pady=10, sticky=tk.W
        )
        name_entry = tb.Entry(main_frame, width=25)
        name_entry.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W + tk.E)
        tb.Label(main_frame, text="主DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=2, column=0, padx=5, pady=5, sticky=tk.W
        )
        primary_entry = tb.Entry(main_frame, width=25)
        primary_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W + tk.E)
        tb.Label(main_frame, text="备用DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=3, column=0, padx=5, pady=5, sticky=tk.W
        )
        secondary_entry = tb.Entry(main_frame, width=25)
        secondary_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W + tk.E)
        main_frame.columnconfigure(1, weight=1)

        def add_and_close():
            selected_category = category_var.get()
            name = name_entry.get().strip()
            primary = primary_entry.get().strip()
            secondary = secondary_entry.get().strip()
            if not name or not primary:
                messagebox.showerror("错误", "名称和主DNS不能为空")
                return
            # 如果选择的分类与当前分类不同，需要切换到该分类
            if selected_category != self.current_category:
                self.current_category = selected_category
                self.category_combo.set(selected_category)
                self.load_category_dns(selected_category)
            self.dns_servers.append(
                {
                    "name": name,
                    "primary": primary,
                    "secondary": secondary,
                }
            )
            self.update_treeview()
            self.auto_save_current_category()  # 自动保存
            dns_dialog.destroy()

        # 创建按钮
        button_frame = tb.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=15)
        tb.Button(
            button_frame, text="添加", command=add_and_close, bootstyle=SUCCESS
        ).pack(side=tk.LEFT, padx=10)
        tb.Button(
            button_frame, text="取消", command=dns_dialog.destroy, bootstyle=SECONDARY
        ).pack(side=tk.LEFT, padx=10)
        name_entry.focus()

    def remove_dns(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要删除的DNS服务器")
            return
        if messagebox.askyesno("确认", "确定要删除选中的DNS服务器吗?"):
            for item in selected_items:
                name = self.tree.item(item, "values")[0]
                self.dns_servers = [s for s in self.dns_servers if s["name"] != name]
            self.update_treeview()
            self.auto_save_current_category()  # 自动保存

    def update_treeview(self):
        # 清空当前所有项
        for item in self.tree.get_children():
            self.tree.delete(item)
        # 填充数据
        for index, server in enumerate(self.dns_servers):
            name = server["name"]
            primary = server["primary"]
            secondary = server["secondary"]
            # 获取测试结果
            result = self.test_results.get(name, {})
            primary_latency = result.get("primary_latency", "未测试")
            secondary_latency = result.get("secondary_latency", "未测试")
            status = result.get("status", "")
            # 格式化延迟显示 (主DNS | 备用DNS)
            if primary_latency != "未测试" and secondary_latency != "未测试":
                if primary_latency == float("inf"):
                    primary_str = "∞"
                else:
                    primary_str = f"{primary_latency:.1f}"
                if secondary_latency == float("inf"):
                    secondary_str = "∞"
                else:
                    secondary_str = f"{secondary_latency:.1f}"
                latency_display = f"{primary_str} | {secondary_str}"
            elif primary_latency != "未测试":
                if primary_latency == float("inf"):
                    latency_display = "∞ | -"
                else:
                    latency_display = f"{primary_latency:.1f} | -"
            else:
                latency_display = "未测试"
            # 设置状态颜色
            if status == "成功":
                status_display = "✔ 成功"
                tag = "success"
            elif status == "失败":
                status_display = "✘ 失败"
                tag = "error"
            else:
                status_display = ""
                tag = ""
            # 设置交替行颜色和状态标签
            row_type = "even" if index % 2 == 0 else "odd"
            # 根据状态和行类型组合标签
            if tag == "success":
                final_tag = f"success_{row_type}"
            elif tag == "error":
                final_tag = f"error_{row_type}"
            else:
                final_tag = f"{row_type}row"
            self.tree.insert(
                "",
                tk.END,
                values=(name, primary, secondary, latency_display, status_display),
                tags=(final_tag,),
            )

    def start_test(self):
        if not self.dns_servers:
            messagebox.showinfo("提示", "请先添加DNS服务器")
            return
        # 重置测试结果
        self.test_results = {}
        # 启动测试线程
        self.update_status("正在测试DNS服务器...")
        threading.Thread(target=self.run_dns_tests, daemon=True).start()

    def run_dns_tests(self):
        test_domain = "www.baidu.com"  # 测试域名
        total = len(self.dns_servers)
        for i, server in enumerate(self.dns_servers):
            name = server["name"]
            primary = server["primary"]
            secondary = server["secondary"]
            # 更新状态栏
            self.root.after(
                0,
                lambda s=f"正在测试 {name} ({i + 1}/{total})...": self.status_var.set(
                    s
                ),
            )
            # 测试主DNS
            latency_primary, status_primary = self.test_dns(primary, test_domain)
            # 如果主DNS测试失败且有备用DNS，测试备用DNS
            latency_secondary = None
            if status_primary == "失败" and secondary:
                latency_secondary, status_secondary = self.test_dns(
                    secondary, test_domain
                )
            else:
                latency_secondary, status_secondary = latency_primary, status_primary
            # 计算平均延迟（如果两个都成功）
            if status_primary == "成功" and status_secondary == "成功":
                avg_latency = (latency_primary + latency_secondary) / 2
            elif status_primary == "成功":
                avg_latency = latency_primary
            elif status_secondary == "成功":
                avg_latency = latency_secondary
            else:
                avg_latency = float("inf")
            # 保存测试结果
            self.test_results[name] = {
                "latency": round(avg_latency, 2)
                if avg_latency != float("inf")
                else "∞",
                "status": "成功"
                if status_primary == "成功" or status_secondary == "成功"
                else "失败",
                "primary_latency": latency_primary,
                "primary_status": status_primary,
                "secondary_latency": latency_secondary,
                "secondary_status": status_secondary,
            }
            # 更新列表视图
            self.root.after(0, self.update_treeview)
        # 排序DNS服务器（延迟低的在前）
        self.dns_servers.sort(
            key=lambda s: self.test_results.get(s["name"], {}).get(
                "latency", float("inf")
            )
            if self.test_results.get(s["name"], {}).get("status") == "成功"
            else float("inf")
        )
        # 更新列表视图
        self.root.after(0, self.update_treeview)
        self.root.after(0, lambda: self.status_var.set("测试完成"))

    def test_dns(self, dns_server, domain):
        try:
            resolver = dns.resolver.Resolver()
            resolver.nameservers = [dns_server]
            resolver.timeout = 3  # 设置超时时间为3秒（IPv6可能需要更长时间）
            resolver.lifetime = 3  # 设置查询生命周期为3秒
            start_time = time.time()
            answers = resolver.resolve(domain)  # noqa: F841
            end_time = time.time()
            latency = (end_time - start_time) * 1000  # 转换为毫秒
            return latency, "成功"
        except:  # noqa: E722
            return float("inf"), "失败"

    def clear_results(self):
        self.test_results = {}
        self.update_treeview()
        self.update_status("结果已清空")

    def move_to_category(self, target_category):
        """将选中的DNS服务器移动到其他分类"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_notification("请先选择要移动的DNS服务器", WARNING)
            return
        # 获取选中的DNS服务器信息
        servers_to_move = []
        for item in selected_items:
            name = self.tree.item(item, "values")[0]
            for server in self.dns_servers:
                if server["name"] == name:
                    servers_to_move.append(server.copy())
                    break
        if not servers_to_move:
            return
        try:
            # 读取配置文件
            config = configparser.ConfigParser()
            config.read("dns_servers.ini", encoding="utf-8")
            # 确保目标分类存在
            if target_category not in config.sections():
                config.add_section(target_category)
            # 将DNS服务器添加到目标分类
            for server in servers_to_move:
                if server["secondary"]:
                    value = f"{server['primary']},{server['secondary']}"
                else:
                    value = server["primary"]
                config.set(target_category, server["name"], value)
            # 从当前分类中删除
            for server in servers_to_move:
                if config.has_option(self.current_category, server["name"]):
                    config.remove_option(self.current_category, server["name"])
                # 从内存中删除
                self.dns_servers = [
                    s for s in self.dns_servers if s["name"] != server["name"]
                ]
            # 保存配置文件
            with open("dns_servers.ini", "w", encoding="utf-8") as f:
                config.write(f)
            # 更新界面
            self.update_treeview()
            # 显示通知
            count = len(servers_to_move)
            self.show_notification(
                f"已将 {count} 个DNS服务器移动到 {target_category}", SUCCESS
            )
        except Exception as e:
            self.show_notification(f"移动DNS服务器失败: {str(e)}", DANGER)

    def clear_dns_cache(self):
        """使用WMI清理系统DNS缓存"""
        try:
            # 初始化COM环境
            pythoncom.CoInitialize()
            # 连接到WMI服务的网络客户端命名空间
            wmi = win32com.client.GetObject("winmgmts:\\\\.\\root\\StandardCimv2")
            # 获取DNS客户端配置对象
            dns_clients = wmi.ExecQuery("SELECT * FROM MSFT_DNSClientCache")
            if not dns_clients:
                self.show_notification("未找到DNS缓存对象", DANGER)
            # 清理DNS缓存
            for client in dns_clients:
                result = client.ClearCache()
                # 检查结果 (0表示成功)
                if result == 0:
                    self.show_notification("DNS缓存已清理", SUCCESS)
                else:
                    raise Exception(f"WMI执行失败，错误代码: {result[0]}")
            return False, "未执行DNS缓存清理操作"
        except Exception:
            self.show_notification("清理DNS缓存失败", DANGER)
        finally:
            # 释放COM环境
            pythoncom.CoUninitialize()

    def _set_dns_via_wmi(self, adapter_name, dns_servers=None, enable_dhcp=False):
        """使用WMI设置DNS服务器"""
        try:
            pythoncom.CoInitialize()
            # 使用Dispatch创建WMI对象，这样更可靠
            wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            wmi = wmi_service.ConnectServer(".", "root\\cimv2")

            # 首先获取所有启用的网络适配器配置
            configs = wmi.ExecQuery(
                "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"
            )

            # 获取网络适配器信息以匹配名称
            adapters = wmi.ExecQuery(
                "SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True"
            )
            adapter_map = {}
            for adapter in adapters:
                adapter_map[str(adapter.DeviceID)] = {
                    "name": adapter.Name,
                    "net_connection_id": adapter.NetConnectionID,
                }

            # 查找匹配的适配器配置
            target_config = None
            for config in configs:
                adapter_info = adapter_map.get(str(config.Index), {})
                # 优先匹配NetConnectionID（友好名称），然后匹配Name
                friendly_name = adapter_info.get("net_connection_id", "")
                adapter_full_name = adapter_info.get("name", "")

                if (
                    friendly_name == adapter_name
                    or adapter_full_name == adapter_name
                    or adapter_name in friendly_name
                    or adapter_name in adapter_full_name
                ):
                    target_config = config
                    print(f"找到匹配的网络适配器: {friendly_name or adapter_full_name}")
                    break

            if not target_config:
                print(f"未找到网络适配器: {adapter_name}")
                return False

            # 调用WMI方法设置DNS
            try:
                # 检查对象类型和可用方法
                print(f"target_config类型: {type(target_config)}")

                # 尝试使用ExecMethod调用
                if enable_dhcp:
                    print("设置DNS为DHCP自动获取")
                    try:
                        # 创建方法参数对象
                        method_params = target_config.Methods_(
                            "SetDNSServerSearchOrder"
                        ).InParameters.SpawnInstance_()
                        method_params.DNSServerSearchOrder = None
                        result = target_config.ExecMethod_(
                            "SetDNSServerSearchOrder", method_params
                        )
                    except Exception as e:
                        print(f"ExecMethod DHCP调用失败: {e}")
                        # 尝试直接调用
                        result = target_config.ExecMethod_("SetDNSServerSearchOrder")
                else:
                    if dns_servers:
                        dns_list = list(dns_servers)
                        print(f"设置DNS服务器列表: {dns_list}")
                        try:
                            # 创建方法参数对象
                            method_params = target_config.Methods_(
                                "SetDNSServerSearchOrder"
                            ).InParameters.SpawnInstance_()
                            method_params.DNSServerSearchOrder = dns_list
                            result = target_config.ExecMethod_(
                                "SetDNSServerSearchOrder", method_params
                            )
                        except Exception as e:
                            print(f"ExecMethod静态DNS调用失败: {e}")
                            return False
                    else:
                        print("DNS服务器列表为空")
                        return False

                # 处理返回结果
                print(f"WMI调用原始结果: {result}")
                print(f"结果类型: {type(result)}")

                # ExecMethod返回的是一个对象，需要获取ReturnValue属性
                error_code = None
                try:
                    # 打印所有可用属性以便调试
                    print(
                        f"result对象的属性: {[attr for attr in dir(result) if not attr.startswith('_')]}"
                    )

                    # 尝试多种方式获取返回值
                    if hasattr(result, "ReturnValue"):
                        error_code = result.ReturnValue
                        print(f"从ReturnValue获取: {error_code}")
                    elif hasattr(result, "Properties_"):
                        try:
                            # 尝试从Properties中获取ReturnValue
                            return_value_prop = result.Properties_("ReturnValue")
                            if return_value_prop:
                                error_code = return_value_prop.Value
                                print(f"从Properties获取: {error_code}")
                        except:
                            # 尝试遍历所有属性
                            try:
                                for prop in result.Properties_:
                                    print(f"属性: {prop.Name} = {prop.Value}")
                                    if prop.Name == "ReturnValue":
                                        error_code = prop.Value
                                        break
                            except Exception as e2:
                                print(f"遍历属性失败: {e2}")

                    # 如果还是没有获取到，尝试直接访问
                    if error_code is None:
                        try:
                            error_code = int(result)
                        except:
                            error_code = 0  # 假设成功，因为方法调用没有抛出异常
                            print("无法获取错误代码，假设成功")

                except Exception as e:
                    print(f"提取错误代码失败: {e}")
                    error_code = 0  # 假设成功

                print(f"提取的错误代码: {error_code}")

                if error_code == 0:
                    action = "DHCP DNS" if enable_dhcp else "静态DNS"
                    print(f"WMI设置{action}成功")
                    return True
                else:
                    action = "DHCP DNS" if enable_dhcp else "静态DNS"
                    print(f"WMI设置{action}失败，错误代码: {error_code}")
                    return False

            except Exception as e:
                print(f"WMI方法调用失败: {e}")
                print(f"异常类型: {type(e)}")
                import traceback

                traceback.print_exc()
                return False

        except Exception as e:
            print(f"WMI设置DNS失败: {e}")
            import traceback

            traceback.print_exc()
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def reset_to_dhcp(self):
        """恢复DNS为DHCP自动获取"""
        try:
            # 获取当前选择的网络连接
            selected_connection = self.network_var.get()
            if not selected_connection:
                self.show_notification("请先选择网络连接", WARNING)
                return

            # 使用WMI设置
            if self._set_dns_via_wmi(selected_connection, enable_dhcp=True):
                self.show_notification(
                    f"已将 {selected_connection} 的DNS设置为自动获取", SUCCESS
                )
                # 刷新DNS显示
                threading.Thread(target=self._refresh_dns_display, daemon=True).start()
            else:
                print("WMI设置DHCP失败")
                self.show_notification("恢复DNS为自动获取失败", DANGER)

        except Exception as e:
            print(f"恢复DHCP DNS失败: {e}")
            self.show_notification("恢复DNS为自动获取失败", DANGER)

    def show_popup(self, event):
        # 选中鼠标点击的项
        item = self.tree.identify_row(event.y)
        if item:
            # 如果点击的项没有被选中，则选中它
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            # 获取当前选中的项数量
            selected_items = self.tree.selection()
            selected_count = len(selected_items)
            # 清空现有菜单
            self.popup_menu.delete(0, tk.END)
            if selected_count == 1:
                # 单选时显示完整菜单
                self.popup_menu.add_command(label="应用DNS", command=self.apply_dns)
                self.popup_menu.add_separator()
            # 多选和单选都可用的功能
            self.popup_menu.add_command(label="删除选中", command=self.remove_dns)
            self.popup_menu.add_separator()
            # DNS组移动子菜单
            dns_group_menu = tk.Menu(self.popup_menu, tearoff=0)
            for category in self.dns_categories:
                if category != self.current_category:  # 不显示当前分类
                    # 创建一个绑定了category的函数
                    def create_move_command(target_cat):
                        return lambda: self.move_to_category(target_cat)

                    dns_group_menu.add_command(
                        label=f"移动到 {category}",
                        command=create_move_command(category),
                    )
            if dns_group_menu.index(tk.END) is not None:  # 如果有其他分类
                self.popup_menu.add_cascade(label="DNS组", menu=dns_group_menu)
                self.popup_menu.add_separator()
            self.popup_menu.add_command(label="检测当前", command=self.refresh_status)
            # 显示菜单
            self.popup_menu.post(event.x_root, event.y_root)

    def on_double_click(self, event):
        """双击事件处理 - 设置DNS"""
        # 获取双击位置的项目
        item = self.tree.identify_row(event.y)
        if item:
            # 选中该项目
            self.tree.selection_set(item)
            self.tree.focus(item)
            # 调用应用DNS方法
            self.apply_dns()

    def apply_dns(self):
        """应用选中的DNS服务器（包括主DNS和备用DNS）"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择DNS服务器")
            return
        if len(selected_items) > 1:
            messagebox.showinfo("提示", "一次只能设置一个DNS服务器")
            return
        item = selected_items[0]
        name, dns_primary, dns_secondary, _, status = self.tree.item(item, "values")
        if status != "✔ 成功":
            if messagebox.askyesno("警告", f"{name} 测试失败，确定要设置为系统DNS吗?"):
                pass
            else:
                return
        # 获取选中的网络连接
        selected_connection = self.network_var.get()
        if not selected_connection:
            messagebox.showerror("错误", "请先选择网络设备")
            return
        if not dns_primary:
            messagebox.showinfo("提示", f"{name} 的主DNS为空")
            return
        # 设置DNS（包括主DNS和备用DNS）
        result = self.set_network_dns(selected_connection, dns_primary, dns_secondary)
        if result:
            dns_info = f"主DNS: {dns_primary}"
            if dns_secondary and dns_secondary.strip():
                dns_info += f", 备用DNS: {dns_secondary}"
            self.show_notification(f"已应用 {name} 的DNS设置 ({dns_info})", SUCCESS)
            # 刷新DNS显示
            threading.Thread(target=self._refresh_dns_display, daemon=True).start()
        else:
            self.show_notification(f"应用 {name} 的DNS设置失败", DANGER)

    def _refresh_dns_display(self):
        """在后台线程中刷新DNS显示"""
        time.sleep(2)  # 等待DNS设置生效
        self.root.after(0, self.update_ip_display)

    def delete_selected_rows(self):
        """删除选中的行"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要删除的DNS服务器")
            return
        if messagebox.askyesno(
            "确认", f"确定要删除选中的 {len(selected_items)} 个DNS服务器吗?"
        ):
            for item in selected_items:
                name = self.tree.item(item, "values")[0]
                self.dns_servers = [s for s in self.dns_servers if s["name"] != name]
            self.update_treeview()
            self.auto_save_current_category()  # 自动保存
            self.show_notification(f"已删除 {len(selected_items)} 个DNS服务器", SUCCESS)

    def load_network_connections(self):
        """加载网络连接到下拉框"""
        self.update_status("正在加载网络设备...")
        threading.Thread(target=self._load_connections_thread, daemon=True).start()

    def _load_connections_thread(self):
        """在后台线程中加载网络连接"""
        connections = self.get_network_connections()
        self.root.after(0, lambda: self._update_network_combo(connections))

    def _update_network_combo(self, connections):
        """更新网络设备下拉框"""
        self.network_connections = connections
        self.network_combo["values"] = connections
        if connections:
            self.network_combo.set(connections[0])
            self.update_status(f"已加载 {len(connections)} 个网络设备")
        else:
            self.network_combo.set("")
            self.update_status("未找到可用的网络设备")
        # 更新IP地址显示
        self.update_ip_display()

    def update_ip_display(self):
        """更新IP地址、网关和DNS显示"""
        ip = self.get_local_ip()
        gateway = self.get_default_gateway()
        dns_servers = self.get_current_dns_servers()
        self.current_ip = ip
        self.ip_var.set(ip)
        self.gateway_var.set(gateway)
        # 更新DNS显示
        primary_dns = dns_servers[0] if len(dns_servers) > 0 else "无法获取"
        secondary_dns = dns_servers[1] if len(dns_servers) > 1 else "未设置"
        self.primary_dns_var.set(primary_dns)
        self.secondary_dns_var.set(secondary_dns)

    def get_network_connections(self):
        """获取可用的网络连接列表"""
        if platform.system() == "Windows":
            try:
                # 使用WMI获取网络适配器信息
                adapters = self._get_network_adapters_info()
                connections = []
                for adapter in adapters:
                    # 只添加有IP地址的适配器
                    if adapter["ipv4_addresses"] or adapter["ipv6_addresses"]:
                        connections.append(adapter["name"])
                return connections if connections else ["未找到网络连接"]
            except Exception as e:
                print(f"获取网络连接失败: {e}")
                # 备选方案：返回常见的网络连接名称
                return ["以太网", "WLAN", "Wi-Fi"]
        else:
            return []

    def set_network_dns(self, connection_name, primary_dns, secondary_dns=None):
        """设置指定网络连接的DNS服务器"""
        if platform.system() != "Windows":
            self.show_notification("此功能仅支持Windows系统", WARNING)
            return False
        try:
            print(f"开始设置DNS - 连接: {connection_name}")
            print(f"主DNS: {primary_dns}")
            if secondary_dns and secondary_dns.strip():
                print(f"备用DNS: {secondary_dns}")

            # 准备DNS服务器列表
            dns_servers = [primary_dns]
            if secondary_dns and secondary_dns.strip():
                dns_servers.append(secondary_dns)

            # 使用WMI设置DNS
            if self._set_dns_via_wmi(connection_name, dns_servers):
                print("WMI设置DNS成功")
                return True
            else:
                print("WMI设置DNS失败")
                return False

        except Exception as e:
            error_msg = f"设置DNS时发生未知错误: {e}"
            print(error_msg)
            self.show_notification("DNS设置失败", DANGER)
            return False

    def run_as_admin(self, command):
        """以管理员权限运行命令"""
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f"/c {command}", None, 0
            )
            return True
        except Exception as e:
            print(f"以管理员权限运行命令失败: {e}")
            return False

    def refresh_status(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择DNS服务器")
            return
        # 获取选中的服务器
        selected_servers = []
        for item in selected_items:
            name = self.tree.item(item, "values")[0]
            for server in self.dns_servers:
                if server["name"] == name:
                    selected_servers.append(server)
                    break
        if not selected_servers:
            return
        # 启动测试线程
        self.update_status("正在刷新DNS服务器状态...")
        threading.Thread(
            target=lambda: self.refresh_selected_dns(selected_servers), daemon=True
        ).start()

    def refresh_selected_dns(self, servers):
        test_domain = "www.baidu.com"  # 测试域名
        total = len(servers)
        for i, server in enumerate(servers):
            name = server["name"]
            # 更新状态栏
            self.root.after(
                0,
                lambda s=f"正在刷新 {name} ({i + 1}/{total})...": self.status_var.set(
                    s
                ),
            )
            # 测试主DNS
            latency_primary, status_primary = self.test_dns(
                server["primary"], test_domain
            )
            # 如果主DNS测试失败且有备用DNS，测试备用DNS
            latency_secondary = None
            if status_primary == "失败" and server["secondary"]:
                latency_secondary, status_secondary = self.test_dns(
                    server["secondary"], test_domain
                )
            else:
                latency_secondary, status_secondary = latency_primary, status_primary
            # 计算平均延迟（如果两个都成功）
            if status_primary == "成功" and status_secondary == "成功":
                avg_latency = (latency_primary + latency_secondary) / 2
            elif status_primary == "成功":
                avg_latency = latency_primary
            elif status_secondary == "成功":
                avg_latency = latency_secondary
            else:
                avg_latency = float("inf")
            # 保存测试结果
            self.test_results[name] = {
                "latency": round(avg_latency, 2)
                if avg_latency != float("inf")
                else "∞",
                "status": "成功"
                if status_primary == "成功" or status_secondary == "成功"
                else "失败",
                "primary_latency": latency_primary,
                "primary_status": status_primary,
                "secondary_latency": latency_secondary,
                "secondary_status": status_secondary,
            }
            # 更新列表视图
            self.root.after(0, self.update_treeview)
        # 重新排序DNS服务器（延迟低的在前）
        self.dns_servers.sort(
            key=lambda s: self.test_results.get(s["name"], {}).get(
                "latency", float("inf")
            )
            if self.test_results.get(s["name"], {}).get("status") == "成功"
            else float("inf")
        )
        # 更新列表视图
        self.root.after(0, self.update_treeview)
        self.root.after(0, lambda: self.status_var.set("刷新完成"))


def is_admin():
    """检查当前程序是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_as_admin():
    """以管理员权限重新启动程序"""
    try:
        # 获取当前脚本的完整路径
        script_path = os.path.abspath(sys.argv[0])
        # 如果是.py文件，需要通过python.exe运行
        if script_path.endswith(".py"):
            # 获取Python解释器路径
            python_exe = sys.executable
            # 使用ShellExecuteW以管理员权限运行
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                python_exe,
                f'"{script_path}"',
                None,
                1,  # SW_SHOWNORMAL
            )
        else:
            # 如果是.exe文件，直接运行
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                script_path,
                None,
                None,
                1,  # SW_SHOWNORMAL
            )
        if result > 32:
            print("正在以管理员权限重新启动程序...")
            return True
        else:
            print(f"以管理员权限启动失败，错误代码: {result}")
            return False
    except Exception as e:
        print(f"以管理员权限启动失败: {e}")
        return False


if __name__ == "__main__":
    # 强制要求管理员权限运行
    if not is_admin():
        print("检测到程序未以管理员权限运行")
        print("正在以管理员权限重新启动程序...")
        # 直接以管理员权限重启，不显示选择对话框
        if run_as_admin():
            # 成功启动管理员权限版本，退出当前进程
            sys.exit(0)
        else:
            # 如果无法获取管理员权限，显示错误并退出
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "权限错误",
                "DNS测试工具需要管理员权限才能正常运行。\n\n"
                "请右键点击程序，选择'以管理员身份运行'。",
                icon="error",
            )
            root.destroy()
            sys.exit(1)
    else:
        print("程序已以管理员权限运行")

    # 使用默认主题创建窗口
    root = tb.Window(themename="darkly")
    # 设置窗口标题（只显示管理员权限）
    root.title("DNS服务器测试工具（管理员权限）")
    root.iconbitmap("c:/Users/ccy/Documents/vscode/DNS/icon.ico")
    # 创建应用程序实例
    app = DNSTesterApp(root)

    # 在窗口完全创建后设置初始标题栏主题
    def set_initial_title_bar():
        try:
            # 获取当前主题并设置标题栏
            current_theme = app.theme_combo.get().lower()
            is_dark = app._is_dark_theme(current_theme)
            app._set_title_bar_theme(is_dark)
            print(
                f"初始标题栏主题设置完成: {current_theme} ({'深色' if is_dark else '浅色'})"
            )
        except Exception as e:
            print(f"设置初始标题栏主题失败: {e}")

    # 延迟设置标题栏主题，确保窗口完全加载
    root.after(200, set_initial_title_bar)
    root.mainloop()
