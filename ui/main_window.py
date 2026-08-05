"""主窗口模块 - 优化后的DNS测试工具主界面"""

import sys
import tkinter as tk
import threading
import urllib.request
import webbrowser
import ttkbootstrap as tb
from ttkbootstrap.constants import PRIMARY, SECONDARY, SUCCESS, DANGER, WARNING, INFO
from typing import List, Dict, Optional, Callable

from config.settings import AppConfig
from config.config_manager import ConfigManager
from config.models import DNSServer
from core.dns_tester import DNSTester, TestResult
from core.network_manager import NetworkManager
from core.dns_config_handler import DNSConfigHandler
from ui.dialogs import DialogManager
from utils.helpers import (
    center_window,
    is_admin,
    run_as_admin,
    is_ip,
    get_notification_style,
)
from utils.logger import get_logger
from utils.history_manager import HistoryManager

logger = get_logger("MainWindow")


class MainWindow:
    """DNS测试工具主窗口"""

    def __init__(self, root: tk.Tk, on_ready: Callable = None):
        """
        初始化主窗口

        Args:
            root: Tkinter根窗口
            on_ready: 初始化完成回调函数
        """
        self.root = root
        self._on_ready = on_ready
        self._init_window()
        self._init_managers()
        self._init_data()
        self.create_widgets()
        self._load_initial_data()
        logger.info("主窗口初始化完成")

    def _init_window(self):
        """初始化窗口设置"""
        self.root.title(f"{AppConfig.WINDOW_TITLE}（管理员权限）")
        self.root.minsize(*AppConfig.WINDOW_SIZE)
        # 先更新窗口以确保几何属性生效，然后再居中
        self.root.update_idletasks()
        center_window(self.root, *AppConfig.WINDOW_SIZE)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _init_managers(self):
        """初始化管理器"""
        self.config_manager = ConfigManager()
        self.dns_config_handler = DNSConfigHandler(self.config_manager)
        self.network_manager = NetworkManager()
        self.dns_tester = DNSTester(progress_callback=self._on_test_progress)
        self.dialog_manager = DialogManager(self.root)
        self.history_manager = HistoryManager()
        logger.info("所有管理器初始化完成")

    def _init_data(self):
        """初始化数据结构"""
        self.dns_servers: List[DNSServer] = []
        self.test_results: Dict[str, TestResult] = {}
        self.network_connections: List[str] = []
        self.dns_categories: List[str] = []
        self.current_category = "Ipv4_默认"
        self.current_ip = ""
        self._original_status = ""

    def _safe_after(self, delay: int, callback: Callable):
        """
        线程安全的 after() 方法

        在后台线程中调用 root.after() 时，如果主循环已退出会抛出异常。
        此方法会捕获异常并忽略，避免程序崩溃。

        Args:
            delay: 延迟时间（毫秒）
            callback: 回调函数
        """
        try:
            self.root.after(delay, callback)
        except (tk.TclError, RuntimeError):
            # 窗口已被销毁或主循环已退出，忽略更新
            pass

    def _load_initial_data(self):
        """加载初始数据"""
        self.load_default_config()
        self.load_theme_preference()
        self.load_network_connections()
        # update_ip_display 会在网络连接加载完成后自动调用

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

        # 左侧：本地网络信息
        left_frame = tb.Frame(network_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        info_configs = [
            ("当前IP:", "ip_var", SUCCESS),
            ("默认网关:", "gateway_var", INFO),
            ("主DNS:", "primary_dns_var", PRIMARY),
            ("备用DNS:", "secondary_dns_var", WARNING),
        ]

        for label_text, var_name, style in info_configs:
            tb.Label(left_frame, text=label_text).pack(side=tk.LEFT, padx=(0, 5))
            var = tk.StringVar(value="获取中...")
            setattr(self, var_name, var)
            label = tb.Label(left_frame, textvariable=var, bootstyle=style)
            label.pack(side=tk.LEFT, padx=(0, 5))

        # 右侧：IPv6 连通性检测（点击链接测试）
        right_frame = tb.Frame(network_frame)
        right_frame.pack(side=tk.RIGHT, padx=(10, 0))

        self.ext_ipv6_var = tk.StringVar(value="IPv6: 检测中...")
        ipv6_label = tb.Label(
            right_frame,
            textvariable=self.ext_ipv6_var,
            font=(AppConfig.FONT_FAMILY, 9),
            cursor="hand2",
            foreground="#007bff",
        )
        ipv6_label.pack(side=tk.LEFT, padx=(0, 5))
        ipv6_label.bind("<Button-1>", lambda e: webbrowser.open("https://v6test.ipgg.cn/"))

    def _create_control_frame(self):
        """创建控制区域"""
        category_frame = tb.Frame(self.root)
        category_frame.pack(fill=tk.X, padx=5, pady=5)

        self._create_network_device_controls(category_frame)
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
        tb.Button(parent, text="刷新", command=self.load_network_connections, bootstyle=INFO).pack(
            side=tk.LEFT, padx=(0, 15)
        )

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

        for i in range(len(buttons)):
            button_frame.grid_columnconfigure(i, weight=1)

        for i, (text, command, style) in enumerate(buttons):
            tb.Button(button_frame, text=text, command=command, bootstyle=style).grid(
                row=0, column=i, padx=2, sticky="ew"
            )

    def _create_status_frame(self):
        """创建状态栏"""
        status_frame = tb.Frame(self.root, bootstyle=SECONDARY)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        self.status_var = tk.StringVar(value="就绪 - 管理员权限")
        self.status_bar = tb.Label(status_frame, textvariable=self.status_var, bootstyle=SECONDARY)
        self.status_bar.pack(side=tk.LEFT, padx=10, pady=5)

        # GitHub链接
        github_label = tb.Label(status_frame, text="⭐ GitHub", bootstyle="INFO", cursor="hand2")
        github_label.pack(side=tk.RIGHT, padx=10, pady=5)
        github_label.bind("<Button-1>", self.open_github)

    def _create_treeview_frame(self):
        """创建左右分栏布局：左侧分类列表 + 右侧DNS服务器列表"""
        main_frame = tb.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # 左侧：分类列表
        self._create_category_list(main_frame)

        # 右侧：DNS服务器列表
        self._create_server_treeview(main_frame)

    def _create_category_list(self, parent):
        """创建左侧分类列表"""
        left_frame = tb.Frame(parent)
        left_frame.grid(row=0, column=0, sticky="ns", padx=(0, 5))

        # 标题
        tb.Label(left_frame, text="DNS类别", font=(AppConfig.FONT_FAMILY, 10, "bold")).pack(
            fill=tk.X, pady=(0, 5)
        )

        # 列表框
        list_container = tb.Frame(left_frame)
        list_container.pack(fill=tk.BOTH, expand=True)

        self.category_listbox = tk.Listbox(
            list_container,
            font=(AppConfig.FONT_FAMILY, 10),
            width=12,
            height=15,
            activestyle="none",
            borderwidth=1,
            relief="solid",
            exportselection=False,
        )
        self.category_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tb.Scrollbar(
            list_container, orient=tk.VERTICAL, command=self.category_listbox.yview
        )
        self.category_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 右键菜单
        self._create_category_context_menu()

        # 拖放绑定
        self._setup_category_drag_drop()

        # 点击切换分类
        self.category_listbox.bind("<<ListboxSelect>>", self.on_category_changed)

    def _create_category_context_menu(self):
        """创建分类右键菜单"""
        self.category_menu = tk.Menu(self.category_listbox, tearoff=0)
        self.category_menu.add_command(label="添加分类", command=self._add_category)
        self.category_menu.add_command(label="重命名", command=self._rename_category)
        self.category_menu.add_separator()
        self.category_menu.add_command(label="上移", command=self._move_category_up)
        self.category_menu.add_command(label="下移", command=self._move_category_down)
        self.category_menu.add_separator()
        self.category_menu.add_command(label="删除分类", command=self._delete_category)

        self.category_listbox.bind("<Button-3>", self._on_category_right_click)

    def _on_category_right_click(self, event):
        """分类右键菜单"""
        idx = self.category_listbox.nearest(event.y)
        if idx >= 0:
            self.category_listbox.selection_clear(0, tk.END)
            self.category_listbox.selection_set(idx)
            self.category_listbox.see(idx)
        self.category_menu.post(event.x_root, event.y_root)

    def _setup_category_drag_drop(self):
        """设置分类列表的拖放移动"""
        self._drag_start_index = None

        self.category_listbox.bind("<ButtonPress-1>", self._on_drag_start)
        self.category_listbox.bind("<B1-Motion>", self._on_drag_motion)
        self.category_listbox.bind("<ButtonRelease-1>", self._on_drag_end)

    def _on_drag_start(self, event):
        """拖放开始"""
        # 检查点击位置是否在实际项目上
        idx = self.category_listbox.nearest(event.y)
        # 检查点击位置是否在有效范围内
        if idx >= 0:
            bbox = self.category_listbox.bbox(idx)
            if bbox is None:
                return
            # 检查点击是否在项目的垂直范围内
            item_top = bbox[1]
            item_bottom = item_top + bbox[3]
            if event.y < item_top or event.y > item_bottom:
                return
        self._drag_start_index = idx

    def _on_drag_motion(self, event):
        """拖放移动中"""
        if self._drag_start_index is None:
            return
        target_idx = self.category_listbox.nearest(event.y)
        if target_idx != self._drag_start_index and 0 <= target_idx < len(self.dns_categories):
            # 交换位置
            category = self.dns_categories[self._drag_start_index]
            self.dns_categories.pop(self._drag_start_index)
            self.dns_categories.insert(target_idx, category)
            self._refresh_category_list()
            self.category_listbox.selection_set(target_idx)
            self._drag_start_index = target_idx

    def _on_drag_end(self, event):
        """拖放结束"""
        if self._drag_start_index is not None:
            self.dns_config_handler.update_category_order(self.dns_categories)
            self._drag_start_index = None

    def _refresh_category_list(self):
        """刷新分类列表显示"""
        self.category_listbox.delete(0, tk.END)
        for category in self.dns_categories:
            self.category_listbox.insert(tk.END, category)

    def _add_category(self):
        """添加分类"""
        new_name = "新分类"
        counter = 1
        while new_name in self.dns_categories:
            new_name = f"新分类{counter}"
            counter += 1

        if self.dns_config_handler.add_category(new_name):
            self.dns_categories.append(new_name)
            self._refresh_category_list()
            self.dns_config_handler.update_category_order(self.dns_categories)
            # 选中新分类并启动编辑
            idx = len(self.dns_categories) - 1
            self.category_listbox.selection_set(idx)
            self.category_listbox.see(idx)
            self.show_notification(f"已添加分类: {new_name}", SUCCESS)

    def _rename_category(self):
        """重命名分类"""
        selection = self.category_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        old_name = self.dns_categories[idx]

        if old_name in ("Ipv4_默认", "Ipv6_默认"):
            self.dialog_manager.show_warning_dialog("警告", "不能重命名默认分类")
            return

        # 内联编辑
        bbox = self.category_listbox.bbox(idx)
        if bbox is None:
            return
        x, y, width, height = bbox

        entry = tk.Entry(
            self.category_listbox,
            font=(AppConfig.FONT_FAMILY, 10),
            bd=0,
            highlightthickness=1,
            highlightbackground="#007bff",
        )
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        listbox_width = self.category_listbox.winfo_width()
        entry.place(x=0, y=y, width=listbox_width, height=height)
        entry.focus_set()

        def save_rename(_event=None):
            new_name = entry.get().strip()
            entry.destroy()
            if new_name and new_name != old_name:
                from utils.helpers import sanitize_input

                new_name = sanitize_input(new_name)
                if new_name and new_name not in self.dns_categories:
                    if self.dns_config_handler.rename_category(old_name, new_name):
                        self.dns_categories[idx] = new_name
                        self._refresh_category_list()
                        self.category_listbox.selection_set(idx)
                        self.dns_config_handler.update_category_order(self.dns_categories)
                        self.show_notification(f"已重命名: {old_name} -> {new_name}", SUCCESS)

        entry.bind("<Return>", save_rename)
        entry.bind("<Escape>", lambda _: entry.destroy())
        entry.bind("<FocusOut>", save_rename)

    def _move_category_up(self):
        """上移分类"""
        selection = self.category_listbox.curselection()
        if not selection or selection[0] == 0:
            return
        idx = selection[0]
        category = self.dns_categories[idx]
        self.dns_categories.pop(idx)
        self.dns_categories.insert(idx - 1, category)
        self._refresh_category_list()
        self.category_listbox.selection_set(idx - 1)
        self.category_listbox.see(idx - 1)
        self.dns_config_handler.update_category_order(self.dns_categories)

    def _move_category_down(self):
        """下移分类"""
        selection = self.category_listbox.curselection()
        if not selection or selection[0] >= len(self.dns_categories) - 1:
            return
        idx = selection[0]
        category = self.dns_categories[idx]
        self.dns_categories.pop(idx)
        self.dns_categories.insert(idx + 1, category)
        self._refresh_category_list()
        self.category_listbox.selection_set(idx + 1)
        self.category_listbox.see(idx + 1)
        self.dns_config_handler.update_category_order(self.dns_categories)

    def _delete_category(self):
        """删除分类"""
        selection = self.category_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        category = self.dns_categories[idx]

        if category in ("Ipv4_默认", "Ipv6_默认"):
            self.dialog_manager.show_warning_dialog("警告", "不能删除默认分类")
            return

        if category == self.current_category:
            self.dialog_manager.show_warning_dialog("警告", "不能删除当前正在使用的分类")
            return

        if self.dns_config_handler.delete_category(category):
            self.dns_categories.pop(idx)
            self._refresh_category_list()
            self.dns_config_handler.update_category_order(self.dns_categories)
            self.show_notification(f"已删除分类: {category}", WARNING)

    def _create_server_treeview(self, parent):
        """创建右侧DNS服务器列表"""
        right_frame = tb.Frame(parent)
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)

        columns = ("name", "primary", "secondary", "latency", "status")
        self.tree = tb.Treeview(
            right_frame, columns=columns, show="headings", bootstyle=INFO, height=30
        )

        self._setup_treeview_style()
        self._setup_treeview_columns()
        self._setup_treeview_scrollbar(right_frame)
        self._setup_treeview_events()

    def _setup_treeview_style(self):
        """设置Treeview样式"""
        style = tb.Style()
        style.configure(
            "Custom.Treeview",
            relief="solid",
            borderwidth=1,
            rowheight=AppConfig.TREEVIEW_ROW_HEIGHT,
        )
        style.configure("Custom.Treeview.Heading", relief="solid", borderwidth=1)
        self.tree.configure(style="Custom.Treeview")

        # 根据当前主题选择隔行颜色
        self._apply_treeview_tag_colors()

    def _apply_treeview_tag_colors(self):
        """根据当前主题（深色/浅色）设置tag颜色"""
        style = tb.Style()
        current_theme = style.theme_use()
        is_dark = current_theme in AppConfig.DARK_THEMES

        if is_dark:
            even_bg = "#2d2d2d"
            odd_bg = "#383838"
        else:
            even_bg = "#f0f0f0"
            odd_bg = "#ffffff"

        self.tree.tag_configure("evenrow", background=even_bg)
        self.tree.tag_configure("oddrow", background=odd_bg)
        self.tree.tag_configure("success", foreground="#28a745")
        self.tree.tag_configure("error", foreground="#dc3545")

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
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

    def _setup_treeview_events(self):
        """设置Treeview事件"""
        self.popup_menu = tk.Menu(self.root, tearoff=0)
        self.tree.bind("<Button-3>", self.show_popup)
        self.tree.bind("<Double-1>", self.on_double_click)

    # 网络信息相关方法
    def load_network_connections(self):
        """加载网络连接"""
        self.update_status("正在加载网络设备...")
        threading.Thread(target=self._load_connections_thread, daemon=True).start()

    def _load_connections_thread(self):
        """在后台线程中加载网络连接"""
        connections = self.network_manager.get_network_connections()
        self._safe_after(0, lambda: self._update_network_combo(connections))

    def _update_network_combo(self, connections: List[str]):
        """更新网络设备下拉框"""
        self.network_connections = connections
        self.network_combo["values"] = connections
        if connections:
            self.network_combo.set(connections[0])
            self.update_status(f"已加载 {len(connections)} 个网络设备")
        else:
            self.network_combo.set("")
            self.update_status("未找到可用的网络设备")
        self.update_ip_display()

        # 首次加载完成后触发回调
        if self._on_ready:
            self._on_ready()
            self._on_ready = None  # 只触发一次

    def update_ip_display(self):
        """更新IP地址、网关和DNS显示"""
        ip = self.network_manager.get_local_ip()
        gateway = self.network_manager.get_default_gateway()
        dns_servers = self.network_manager.get_current_dns_servers()

        self.current_ip = ip
        self.ip_var.set(ip)
        self.gateway_var.set(gateway)

        primary_dns = dns_servers[0] if len(dns_servers) > 0 else "无法获取"
        secondary_dns = dns_servers[1] if len(dns_servers) > 1 else "未设置"
        self.primary_dns_var.set(primary_dns)
        self.secondary_dns_var.set(secondary_dns)

        # 网络信息加载完成后自动检测外网IP
        self._query_external_ip_info()

    def _query_external_ip_info(self):
        """在后台线程中查询外网IPv4/IPv6地址和连接类型"""
        threading.Thread(target=self._query_external_ip_thread, daemon=True).start()

    def _query_external_ip_thread(self):
        """后台线程检测IPv6连通性"""
        try:
            req = urllib.request.Request("https://v6.ipgg.cn", headers={"User-Agent": "curl/8.0"})
            with urllib.request.urlopen(req, timeout=5) as resp:
                resp.read().decode("utf-8", errors="ignore").strip()
            self._safe_after(0, lambda: self.ext_ipv6_var.set("IPv6: 支持"))
        except Exception:
            self._safe_after(0, lambda: self.ext_ipv6_var.set("IPv6: 不支持"))

    # 配置相关方法
    def load_default_config(self):
        """加载默认配置"""
        if not self.config_manager.load_config():
            self.dns_config_handler.create_default_config()
            self.config_manager.load_config()

        self.dns_categories = self.dns_config_handler.load_categories()
        if not self.dns_categories:
            self.dns_categories = ["Ipv4_默认", "Ipv6_默认"]

        # 加载上次选择的分类
        last_category = self.dns_config_handler.get_last_category()
        if last_category and last_category in self.dns_categories:
            self.current_category = last_category
        else:
            self.current_category = self.dns_categories[0]

        self._refresh_category_list()
        if self.current_category in self.dns_categories:
            idx = self.dns_categories.index(self.current_category)
            self.category_listbox.selection_set(idx)
            self.category_listbox.see(idx)

        self.dns_servers = self.dns_config_handler.load_category_servers(self.current_category)
        self.update_treeview()
        self.show_notification("已加载默认DNS配置", INFO)

    def load_theme_preference(self):
        """加载主题偏好"""
        theme_name = self.dns_config_handler.get_theme()
        theme_display = theme_name.capitalize()
        self.theme_combo.set(theme_display)
        # 真正应用主题（不仅是设combo）
        style = tb.Style()
        style.theme_use(theme_name)
        # 主题切换会重置样式，重新应用行高
        style.configure(
            "Custom.Treeview",
            relief="solid",
            borderwidth=1,
            rowheight=AppConfig.TREEVIEW_ROW_HEIGHT,
        )
        style.configure("Custom.Treeview.Heading", relief="solid", borderwidth=1)
        self.tree.configure(style="Custom.Treeview")
        is_dark = theme_name in AppConfig.DARK_THEMES
        self.root.after(100, lambda: self._set_title_bar_theme(is_dark))
        # 确保列表颜色与已加载的主题同步
        self._apply_treeview_tag_colors()

    # DNS测试相关方法
    def start_test(self):
        """开始DNS测试"""
        if not self.dns_servers:
            self.dialog_manager.show_info_dialog("提示", "没有可测试的DNS服务器")
            return

        # 检查代理设置
        from utils.proxy_detector import ProxyDetector

        is_proxy_enabled, proxy_info = ProxyDetector.check_system_proxy()

        if is_proxy_enabled:
            warning_msg = ProxyDetector.get_proxy_warning_message(proxy_info)
            self.dialog_manager.show_confirm_dialog(
                "代理警告",
                f"检测到系统代理已启用，可能影响测试结果\n\n{warning_msg}\n\n代理可能会影响DNS测试结果的准确性\n\n是否继续？",
                on_confirm=lambda: self._start_dns_test(),
            )
        else:
            self._start_dns_test()

    def _start_dns_test(self):
        """执行DNS测试"""
        self.test_results = {}
        self.update_status("正在测试DNS服务器...")

        # 准备数据
        servers_data = [
            {
                "name": server.name,
                "primary": server.primary,
                "secondary": server.secondary,
            }
            for server in self.dns_servers
        ]

        # 启动测试
        threading.Thread(
            target=lambda: self._run_dns_tests_parallel(servers_data),
            daemon=True,
        ).start()

    def _run_dns_tests_parallel(self, servers_data: List[Dict]):
        """运行并行DNS测试"""
        try:
            self.test_results = self.dns_tester.test_dns_servers_parallel(servers_data)

            # 保存到历史记录
            for name, result in self.test_results.items():
                server = next((s for s in self.dns_servers if s.name == name), None)
                if server:
                    self.history_manager.add_record(
                        dns_name=name,
                        dns_primary=server.primary,
                        dns_secondary=server.secondary,
                        latency=result.avg_latency,
                        status=result.status,
                        category=self.current_category,
                    )

            # 排序DNS服务器
            self.dns_servers.sort(
                key=lambda s: (
                    self.test_results.get(s.name).avg_latency
                    if self.test_results.get(s.name)
                    and (
                        self.test_results.get(s.name).primary_status == "成功"
                        or self.test_results.get(s.name).secondary_status == "成功"
                    )
                    else float("inf")
                )
            )

            self._safe_after(0, self.update_treeview)
            self._safe_after(0, lambda: self.update_status("测试完成"))
            self.show_notification("DNS测试完成", SUCCESS)

        except Exception as e:
            logger.error(f"DNS测试失败: {e}")
            self._safe_after(0, lambda: self.update_status("测试失败"))
            self.show_notification("DNS测试失败", DANGER)

    def _on_test_progress(self, server_name: str, current: int, total: int):
        """测试进度回调"""
        self._safe_after(
            0,
            lambda: self.update_status(f"正在测试 {server_name} ({current}/{total})..."),
        )

    # DNS设置相关方法
    def apply_dns(self):
        """应用选中的DNS服务器"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.dialog_manager.show_info_dialog("提示", "请先选择DNS服务器")
            return
        if len(selected_items) > 1:
            self.dialog_manager.show_info_dialog("提示", "一次只能设置一个DNS服务器")
            return

        item = selected_items[0]
        values = self.tree.item(item, "values")
        name, dns_primary, dns_secondary, _, status = values

        if status != "✔ 成功":
            if not self.dialog_manager.show_confirm_dialog(
                "警告", f"{name} 测试失败，确定要设置为系统DNS吗?"
            ):
                return

        selected_connection = self.network_var.get()
        if not selected_connection:
            self.dialog_manager.show_error_dialog("错误", "请先选择网络设备")
            return

        dns_servers = [dns_primary]
        if dns_secondary and dns_secondary.strip():
            dns_servers.append(dns_secondary)

        result = self.network_manager.set_dns_servers(selected_connection, dns_servers)

        if result:
            dns_info = f"主DNS: {dns_primary}"
            if dns_secondary and dns_secondary.strip():
                dns_info += f", 备用DNS: {dns_secondary}"
            self.show_notification(f"已应用 {name} 的DNS设置 ({dns_info})", SUCCESS)
            threading.Thread(target=self._refresh_dns_display, daemon=True).start()
        else:
            self.show_notification(f"应用 {name} 的DNS设置失败", DANGER)

    def reset_to_dhcp(self):
        """恢复DNS为DHCP自动获取"""
        selected_connection = self.network_var.get()
        if not selected_connection:
            self.show_notification("请先选择网络连接", WARNING)
            return

        try:
            if self.network_manager.set_dns_servers(selected_connection, [], enable_dhcp=True):
                self.show_notification(f"已将 {selected_connection} 的DNS设置为自动获取", SUCCESS)
                threading.Thread(target=self._refresh_dns_display, daemon=True).start()
            else:
                logger.error(f"重置DNS为DHCP失败，适配器: {selected_connection}")
                self.show_notification("恢复DNS为自动获取失败", DANGER)
        except Exception as e:
            logger.error(f"重置DNS异常: {e}")
            self.show_notification(f"恢复DNS失败: {e}", DANGER)

    def _refresh_dns_display(self):
        """在后台线程中刷新DNS显示"""
        import time

        time.sleep(2)
        self._safe_after(0, self.update_ip_display)

    def clear_dns_cache(self):
        """清理DNS缓存"""
        if self.network_manager.clear_dns_cache():
            self.show_notification("DNS缓存已清理", SUCCESS)
        else:
            self.show_notification("清理DNS缓存失败", DANGER)

    # DNS管理相关方法
    def add_dns(self):
        """添加DNS服务器"""
        self.dialog_manager.show_add_dns_dialog(
            categories=self.dns_categories,
            current_category=self.current_category,
            on_add=self._on_dns_added,
        )

    def _on_dns_added(self, category: str, name: str, primary: str, secondary: str):
        """DNS添加回调"""
        from utils.helpers import validate_dns_ip

        if not validate_dns_ip(primary):
            self.dialog_manager.show_error_dialog("错误", "无效的主DNS地址")
            return

        if secondary and not validate_dns_ip(secondary):
            self.dialog_manager.show_error_dialog("错误", "无效的备用DNS地址")
            return

        # 如果选择的分类与当前分类不同，切换到该分类
        if category != self.current_category:
            self.current_category = category
            if category in self.dns_categories:
                idx = self.dns_categories.index(category)
                self.category_listbox.selection_set(idx)
                self.category_listbox.see(idx)
            self.load_category_dns(category)

        self.dns_servers.append(DNSServer(name=name, primary=primary, secondary=secondary))
        self.update_treeview()
        self._save_current_category()

    def edit_dns(self):
        """编辑选中的DNS服务器"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.dialog_manager.show_info_dialog("提示", "请先选择要编辑的DNS服务器")
            return

        item = selected_items[0]
        values = self.tree.item(item, "values")
        name, primary, secondary = values[0], values[1], values[2]

        self.dialog_manager.show_edit_dns_dialog(
            name=name,
            primary=primary,
            secondary=secondary,
            on_save=self._on_dns_edited,
        )

    def _on_dns_edited(self, name: str, primary: str, secondary: str):
        """DNS编辑回调"""
        from utils.helpers import validate_dns_ip

        if not validate_dns_ip(primary):
            self.dialog_manager.show_error_dialog("错误", "无效的主DNS地址")
            return

        if secondary and not validate_dns_ip(secondary):
            self.dialog_manager.show_error_dialog("错误", "无效的备用DNS地址")
            return

        selected_items = self.tree.selection()
        if not selected_items:
            return

        item = selected_items[0]
        old_name = self.tree.item(item, "values")[0]

        # 更新 dns_servers 列表
        for server in self.dns_servers:
            if server.name == old_name:
                server.name = name
                server.primary = primary
                server.secondary = secondary
                break

        self.update_treeview()
        self._save_current_category()
        self.show_notification(f"已更新DNS服务器: {name}", SUCCESS)

    def remove_dns(self):
        """删除选中的DNS服务器"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.dialog_manager.show_info_dialog("提示", "请先选择要删除的DNS服务器")
            return

        if self.dialog_manager.show_confirm_dialog(
            "确认", f"确定要删除选中的 {len(selected_items)} 个DNS服务器吗?"
        ):
            for item in selected_items:
                name = self.tree.item(item, "values")[0]
                self.dns_servers = [s for s in self.dns_servers if s.name != name]
            self.update_treeview()
            self._save_current_category()
            self.show_notification(f"已删除 {len(selected_items)} 个DNS服务器", SUCCESS)

    # UI更新相关方法
    def update_treeview(self):
        """更新列表视图"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for index, server in enumerate(self.dns_servers):
            result = self.test_results.get(server.name)
            if result:
                # TestResult对象
                primary_latency = result.primary_latency
                secondary_latency = result.secondary_latency
                # 计算状态
                if result.primary_status == "成功" or result.secondary_status == "成功":
                    status = "成功"
                else:
                    status = "失败"
            else:
                # 没有测试结果
                primary_latency = "未测试"
                secondary_latency = "未测试"
                status = ""

            # 格式化延迟显示
            def format_latency(value):
                """格式化延迟值"""
                if value is None or value == "未测试":
                    return "-"
                if value == float("inf"):
                    return "∞"
                return f"{value:.1f}"

            if primary_latency != "未测试" and secondary_latency != "未测试":
                primary_str = format_latency(primary_latency)
                secondary_str = format_latency(secondary_latency)
                latency_display = f"{primary_str} | {secondary_str}"
            elif primary_latency != "未测试":
                primary_str = format_latency(primary_latency)
                latency_display = f"{primary_str} | -"
            else:
                latency_display = "未测试"

            # 设置状态显示
            if status == "成功":
                status_display = "✔ 成功"
                status_tag = "success"
            elif status == "失败":
                status_display = "✘ 失败"
                status_tag = "error"
            else:
                status_display = ""
                status_tag = ""

            row_type = "evenrow" if index % 2 == 0 else "oddrow"
            tags = (status_tag, row_type) if status_tag else (row_type,)

            self.tree.insert(
                "",
                tk.END,
                values=(
                    server.name,
                    server.primary,
                    server.secondary,
                    latency_display,
                    status_display,
                ),
                tags=tags,
            )

    def update_status(self, message: str):
        """更新状态栏"""
        if not hasattr(self, "_original_status") or not self._original_status:
            self.status_var.set(message)
        else:
            self._original_status = message

    def show_notification(self, message: str, style=INFO):
        """显示通知"""
        if not hasattr(self, "_original_status"):
            self._original_status = self.status_var.get()

        prefix, color = get_notification_style(style)
        self.status_var.set(f"{prefix}{message}")
        self.status_bar.configure(foreground=color)
        self.root.after(3000, self.hide_notification)

    def hide_notification(self):
        """隐藏通知"""
        if hasattr(self, "_original_status"):
            self.status_var.set(self._original_status)
            delattr(self, "_original_status")
        else:
            self.status_var.set("就绪 - 管理员权限")
        self.status_bar.configure(foreground="")

    # 事件处理
    def on_category_changed(self, event=None):
        """分类改变事件"""
        selection = self.category_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        selected_category = self.dns_categories[idx]
        if selected_category and selected_category != self.current_category:
            # 保存当前分类的测试结果
            if self.test_results:
                if not hasattr(self, "_category_test_results"):
                    self._category_test_results = {}
                self._category_test_results[self.current_category] = self.test_results.copy()

            self.current_category = selected_category

            # 恢复新分类的测试结果
            if (
                hasattr(self, "_category_test_results")
                and selected_category in self._category_test_results
            ):
                self.test_results = self._category_test_results[selected_category].copy()
            else:
                self.test_results = {}

            self.load_category_dns(selected_category)
            self.dns_config_handler.set_last_category(selected_category)

    def load_category_dns(self, category: str):
        """加载指定分类的DNS"""
        self.dns_servers = self.dns_config_handler.load_category_servers(category)
        self.update_treeview()
        self.show_notification(f"已加载 {category} ({len(self.dns_servers)} 个DNS服务器)", SUCCESS)

    def on_theme_changed(self, event=None):
        """主题改变事件"""
        selected_theme = self.theme_var.get()
        if selected_theme:
            try:
                theme_name = selected_theme.lower()
                style = tb.Style()
                style.theme_use(theme_name)

                # 重新应用行高（主题切换会重置样式）
                style.configure(
                    "Custom.Treeview",
                    relief="solid",
                    borderwidth=1,
                    rowheight=AppConfig.TREEVIEW_ROW_HEIGHT,
                )
                style.configure("Custom.Treeview.Heading", relief="solid", borderwidth=1)
                self.tree.configure(style="Custom.Treeview")
                # 重新配置标签颜色，自动适配深色/浅色主题
                self._apply_treeview_tag_colors()

                is_dark = theme_name in AppConfig.DARK_THEMES
                self.root.after(50, lambda: self._set_title_bar_theme(is_dark))

                theme_type = "深色" if is_dark else "浅色"
                self.show_notification(f"已切换到 {selected_theme} 主题 ({theme_type})", SUCCESS)
                self.dns_config_handler.set_theme(theme_name)
            except Exception:
                self.show_notification(f"切换到 {selected_theme} 主题失败", DANGER)
                self.theme_combo.set("Darkly")

    def _set_title_bar_theme(self, is_dark_theme: bool):
        """设置Windows标题栏主题"""
        try:
            import ctypes

            self.root.update_idletasks()

            try:
                hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                if hwnd == 0:
                    hwnd = self.root.winfo_id()
            except Exception:
                return

            try:
                value = 1 if is_dark_theme else 0
                if sys.getwindowsversion().build >= 22000:
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                else:
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 19

                result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(ctypes.c_int(value)),
                    ctypes.sizeof(ctypes.c_int),
                )
                if result == 0:
                    self.root.update()
            except Exception:
                pass
        except Exception as e:
            logger.error(f"设置标题栏主题失败: {e}")

    def show_popup(self, event):
        """显示右键菜单"""
        item = self.tree.identify_row(event.y)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)

            selected_count = len(self.tree.selection())
            self.popup_menu.delete(0, tk.END)

            if selected_count == 1:
                self.popup_menu.add_command(label="应用DNS", command=self.apply_dns)
                self.popup_menu.add_command(label="编辑", command=self.edit_dns)
                self.popup_menu.add_separator()

            self.popup_menu.add_command(label="删除选中", command=self.remove_dns)
            self.popup_menu.add_separator()
            self.popup_menu.add_command(label="检测当前", command=self.refresh_selected)
            self.popup_menu.post(event.x_root, event.y_root)

    def on_double_click(self, event):
        """双击事件 - 编辑DNS服务器"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.edit_dns()

    def refresh_selected(self):
        """刷新选中的DNS服务器"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.dialog_manager.show_info_dialog("提示", "请先选择DNS服务器")
            return

        selected_servers = []
        for item in selected_items:
            name = self.tree.item(item, "values")[0]
            server = next((s for s in self.dns_servers if s.name == name), None)
            if server:
                selected_servers.append(server)

        if selected_servers:
            self.update_status("正在刷新DNS服务器状态...")
            threading.Thread(
                target=lambda: self._refresh_selected_dns(selected_servers),
                daemon=True,
            ).start()

    def _refresh_selected_dns(self, servers: List[DNSServer]):
        """刷新选中的DNS服务器"""
        servers_data = [
            {
                "name": server.name,
                "primary": server.primary,
                "secondary": server.secondary,
            }
            for server in servers
        ]

        results = self.dns_tester.test_dns_servers_parallel(servers_data)
        self.test_results.update(results)

        # 排序
        self.dns_servers.sort(
            key=lambda s: (
                self.test_results.get(s.name).avg_latency
                if self.test_results.get(s.name)
                and (
                    self.test_results.get(s.name).primary_status == "成功"
                    or self.test_results.get(s.name).secondary_status == "成功"
                )
                else float("inf")
            )
        )

        self._safe_after(0, self.update_treeview)
        self._safe_after(0, lambda: self.update_status("刷新完成"))

    def open_github(self, event):
        """打开GitHub链接"""
        try:
            webbrowser.open("https://github.com/pcoof/dns-tester")
            self.show_notification("已在浏览器中打开GitHub页面", INFO)
        except Exception as e:
            logger.error(f"打开GitHub链接失败: {e}")
            self.show_notification("打开GitHub链接失败", WARNING)

    def _save_current_category(self):
        """保存当前分类"""
        self.dns_config_handler.save_category_servers(self.current_category, self.dns_servers)

    def on_closing(self):
        """窗口关闭处理"""
        try:
            logger.info("程序正在关闭，保存配置...")
            self._save_current_category()
            self.dns_config_handler.set_last_category(self.current_category)
            logger.info("配置保存完成，程序退出")
        except Exception as e:
            logger.error(f"关闭时保存配置失败: {e}")
        finally:
            self.root.destroy()
