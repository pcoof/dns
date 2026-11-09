"""对话框管理模块"""

import sys
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import SUCCESS, DANGER, INFO, SECONDARY
from typing import List, Callable

from config.settings import AppConfig
from utils.helpers import sanitize_input
from utils.logger import get_logger

logger = get_logger("DialogManager")


class DialogManager:
    """对话框管理器"""

    def __init__(self, parent: tk.Tk):
        """
        初始化对话框管理器

        Args:
            parent: 父窗口
        """
        self.parent = parent
        logger.info("对话框管理器初始化完成")

    def _apply_title_bar_theme(self, window: tk.Toplevel) -> None:
        """应用与主窗口一致的标题栏主题"""
        try:
            import ctypes

            window.update_idletasks()

            try:
                hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
                if hwnd == 0:
                    hwnd = window.winfo_id()
            except Exception:
                return

            style = tb.Style()
            current_theme = style.theme_use()
            is_dark = current_theme in AppConfig.DARK_THEMES

            try:
                value = 1 if is_dark else 0
                if sys.getwindowsversion().build >= 22000:
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                else:
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 19

                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(ctypes.c_int(value)),
                    ctypes.sizeof(ctypes.c_int),
                )
                window.update()
            except Exception:
                pass
        except Exception:
            pass

    def show_add_dns_dialog(
        self,
        categories: List[str],
        current_category: str,
        on_add: Callable[[str, str, str, str], None],
    ) -> None:
        """
        显示添加DNS服务器对话框

        Args:
            categories: 可用分类列表
            current_category: 当前分类
            on_add: 添加回调函数 (category, name, primary, secondary)
        """
        dns_dialog = tk.Toplevel(self.parent)
        dns_dialog.title("添加DNS服务器")
        dns_dialog.withdraw()  # 隐藏，防止闪烁
        dns_dialog.resizable(False, False)
        dns_dialog.transient(self.parent)
        dns_dialog.grab_set()

        # 计算居中位置
        self.parent.update_idletasks()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        px = self.parent.winfo_x()
        py = self.parent.winfo_y()
        cx = px + (pw - 420) // 2
        cy = py + (ph - 320) // 2
        dns_dialog.geometry(f"420x320+{cx}+{cy}")
        dns_dialog.deiconify()  # 就绪后显示，无闪烁
        self._apply_title_bar_theme(dns_dialog)

        # 主框架
        main_frame = tb.Frame(dns_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 分类选择
        tb.Label(main_frame, text="分类:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=0, column=0, padx=5, pady=8, sticky=tk.W
        )
        category_var = tk.StringVar()
        category_combo = tb.Combobox(
            main_frame,
            textvariable=category_var,
            width=22,
            font=(AppConfig.FONT_FAMILY, 9),
            state="readonly",
        )
        category_combo["values"] = categories
        category_combo.set(current_category)
        category_combo.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W + tk.E)

        # 名称输入
        tb.Label(main_frame, text="名称:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=1, column=0, padx=5, pady=8, sticky=tk.W
        )
        name_entry = tb.Entry(main_frame, width=25)
        name_entry.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W + tk.E)

        # 主DNS输入
        tb.Label(main_frame, text="主DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=2, column=0, padx=5, pady=5, sticky=tk.W
        )
        primary_entry = tb.Entry(main_frame, width=25)
        primary_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W + tk.E)

        # 备用DNS输入
        tb.Label(main_frame, text="备用DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=3, column=0, padx=5, pady=5, sticky=tk.W
        )
        secondary_entry = tb.Entry(main_frame, width=25)
        secondary_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W + tk.E)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

        def add_and_close():
            """添加并关闭对话框"""
            selected_category = category_var.get()
            name = sanitize_input(name_entry.get().strip())
            primary = sanitize_input(primary_entry.get().strip())
            secondary = sanitize_input(secondary_entry.get().strip())

            if not name or not primary:
                messagebox.showerror("错误", "名称和主DNS不能为空")
                return

            on_add(selected_category, name, primary, secondary)
            dns_dialog.destroy()

        # 按钮框架
        button_frame = tb.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(10, 5))

        tb.Button(button_frame, text="添加", command=add_and_close, bootstyle=SUCCESS).pack(
            side=tk.LEFT, padx=15
        )
        tb.Button(button_frame, text="取消", command=dns_dialog.destroy, bootstyle=SECONDARY).pack(
            side=tk.LEFT, padx=15
        )

        # 确保弹窗始终在主窗口之上（override redirect 后 transient 层级可能失效）
        dns_dialog.lift()
        dns_dialog.attributes("-topmost", True)
        dns_dialog.focus_force()

        name_entry.focus()

    def show_category_manager_dialog(
        self,
        categories: List[str],
        on_add: Callable[[str], None],
        on_delete: Callable[[str], None],
        on_rename: Callable[[str, str], None],
        on_move_up: Callable[[int], None],
        on_move_down: Callable[[int], None],
    ) -> None:
        """
        显示分类管理对话框

        支持：双击编辑、右键菜单、直接添加行

        Args:
            categories: 分类列表（会被直接修改）
            on_add: 添加分类回调
            on_delete: 删除分类回调
            on_rename: 重命名分类回调
            on_move_up: 上移回调
            on_move_down: 下移回调
        """
        manage_dialog = tk.Toplevel(self.parent)
        manage_dialog.title("DNS分类管理")
        manage_dialog.withdraw()
        manage_dialog.resizable(True, True)
        manage_dialog.transient(self.parent)
        manage_dialog.grab_set()

        self.parent.update_idletasks()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        px = self.parent.winfo_x()
        py = self.parent.winfo_y()
        cx = px + (pw - 500) // 2
        cy = py + (ph - 500) // 2
        manage_dialog.geometry(f"500x500+{cx}+{cy}")
        manage_dialog.deiconify()
        self._apply_title_bar_theme(manage_dialog)

        # 主框架
        main_frame = tb.Frame(manage_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题
        tb.Label(main_frame, text="DNS分类管理", font=(AppConfig.FONT_FAMILY, 12, "bold")).pack(
            pady=(0, 10)
        )

        # 列表容器
        list_container = tb.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 列表框
        list_frame = tb.Frame(list_container)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        category_listbox = tk.Listbox(list_frame, font=(AppConfig.FONT_FAMILY, 10))
        category_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tb.Scrollbar(list_frame, orient=tk.VERTICAL, command=category_listbox.yview)
        category_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 排序按钮
        sort_frame = tb.Frame(list_container)
        sort_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))

        # --- 内联编辑状态 ---
        edit_entry = [None]  # 用列表方便内部函数修改

        def start_edit(idx: int):
            """在列表框指定行上启动内联编辑"""
            if edit_entry[0] is not None:
                return  # 已有编辑项，忽略

            text = category_listbox.get(idx)
            bbox = category_listbox.bbox(idx)
            if bbox is None:
                return

            x, y, width, height = bbox
            entry = tk.Entry(
                category_listbox,
                font=(AppConfig.FONT_FAMILY, 10),
                bd=0,
                highlightthickness=1,
                highlightbackground="#007bff",
            )
            entry.insert(0, text)
            entry.select_range(0, tk.END)
            # 全宽覆盖列表框
            listbox_width = category_listbox.winfo_width()
            entry.place(x=0, y=y, width=listbox_width, height=height)
            entry.focus_set()
            edit_entry[0] = (entry, idx, text)  # 保存 (entry, 索引, 原始值)

            def on_enter(_event=None):
                finish_edit(save=True)

            def on_escape(_event=None):
                finish_edit(save=False)

            def on_focus_out(_event=None):
                finish_edit(save=True)

            entry.bind("<Return>", on_enter)
            entry.bind("<Escape>", on_escape)
            entry.bind("<FocusOut>", on_focus_out)

        def finish_edit(save: bool = True):
            """结束内联编辑"""
            if edit_entry[0] is None:
                return
            entry, idx, original = edit_entry[0]
            new_text = entry.get().strip()
            entry.destroy()
            edit_entry[0] = None

            if save and new_text and new_text != original:
                new_text = sanitize_input(new_text)
                if new_text and new_text not in categories:
                    on_rename(original, new_text)
                    refresh_list()
                    category_listbox.selection_set(idx)
            elif save and (not new_text or new_text == original):
                # 未修改或为空，恢复
                pass

        def refresh_list():
            """刷新列表"""
            category_listbox.delete(0, tk.END)
            for category in categories:
                category_listbox.insert(tk.END, category)

        def move_up():
            """上移"""
            selection = category_listbox.curselection()
            if not selection or selection[0] == 0:
                return
            idx = selection[0]
            on_move_up(idx)
            refresh_list()
            category_listbox.selection_set(idx - 1)
            category_listbox.see(idx - 1)

        def move_down():
            """下移"""
            selection = category_listbox.curselection()
            if not selection or selection[0] == len(categories) - 1:
                return
            idx = selection[0]
            on_move_down(idx)
            refresh_list()
            category_listbox.selection_set(idx + 1)
            category_listbox.see(idx + 1)

        tb.Button(sort_frame, text="↑", width=3, command=move_up, bootstyle=INFO).pack(pady=(0, 5))
        tb.Button(sort_frame, text="↓", width=3, command=move_down, bootstyle=INFO).pack()

        # --- 右键菜单 ---
        context_menu = tk.Menu(category_listbox, tearoff=0)
        context_menu.add_command(
            label="添加分类",
            command=lambda: add_category(),
        )
        context_menu.add_command(
            label="编辑分类",
            command=lambda: _context_edit(),
        )
        context_menu.add_separator()
        context_menu.add_command(
            label="删除分类",
            command=lambda: delete_category(),
        )

        def _on_right_click(event):
            """右键菜单"""
            idx = category_listbox.nearest(event.y)
            if idx >= 0:
                category_listbox.selection_clear(0, tk.END)
                category_listbox.selection_set(idx)
            context_menu.post(event.x_root, event.y_root)

        def _context_edit():
            """右键菜单 - 编辑"""
            selection = category_listbox.curselection()
            if selection:
                start_edit(selection[0])

        category_listbox.bind("<Button-3>", _on_right_click)

        # --- 双击编辑 ---
        category_listbox.bind("<Double-1>", lambda e: _on_double_click(e))

        def _on_double_click(event):
            """双击启动内联编辑"""
            idx = category_listbox.nearest(event.y)
            if idx >= 0:
                start_edit(idx)

        # 确保弹窗始终在主窗口之上
        manage_dialog.lift()
        manage_dialog.attributes("-topmost", True)
        manage_dialog.focus_force()

        refresh_list()

        # 按钮框架
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def add_category():
            """添加分类 - 直接插入行并自动编辑"""
            new_name = "新分类"
            # 处理重名
            counter = 1
            while new_name in categories:
                new_name = f"新分类{counter}"
                counter += 1
            on_add(new_name)
            refresh_list()
            # 只有确认已添加到列表后才编辑
            if new_name in categories:
                idx = categories.index(new_name)
                category_listbox.selection_clear(0, tk.END)
                category_listbox.selection_set(idx)
                category_listbox.see(idx)
                manage_dialog.after(50, lambda: start_edit(idx))

        def delete_category():
            """删除分类"""
            selection = category_listbox.curselection()
            if not selection:
                return
            category = category_listbox.get(selection[0])
            on_delete(category)
            refresh_list()

        tb.Button(button_frame, text="添加分类", command=add_category, bootstyle=SUCCESS).pack(
            side=tk.LEFT, padx=5
        )
        tb.Button(
            button_frame, text="删除分类", command=delete_category, bootstyle=DANGER
        ).pack(side=tk.LEFT, padx=5)
        tb.Button(
            button_frame,
            text="关闭",
            command=manage_dialog.destroy,
            bootstyle=SECONDARY,
        ).pack(side=tk.RIGHT, padx=5)

    def show_confirm_dialog(self, title: str, message: str, on_confirm: Callable[[], None]) -> None:
        """
        显示确认对话框

        Args:
            title: 对话框标题
            message: 确认消息
            on_confirm: 确认回调
        """
        if messagebox.askyesno(title, message):
            on_confirm()

    def show_edit_dns_dialog(
        self,
        name: str,
        primary: str,
        secondary: str,
        on_save: Callable[[str, str, str], None],
    ) -> None:
        """
        显示编辑DNS服务器对话框

        Args:
            name: 当前名称
            primary: 当前主DNS
            secondary: 当前备用DNS
            on_save: 保存回调函数 (name, primary, secondary)
        """
        edit_dialog = tk.Toplevel(self.parent)
        edit_dialog.title("编辑DNS服务器")
        edit_dialog.withdraw()
        edit_dialog.resizable(False, False)
        edit_dialog.transient(self.parent)
        edit_dialog.grab_set()

        self.parent.update_idletasks()
        pw = self.parent.winfo_width()
        ph = self.parent.winfo_height()
        px = self.parent.winfo_x()
        py = self.parent.winfo_y()
        cx = px + (pw - 420) // 2
        cy = py + (ph - 260) // 2
        edit_dialog.geometry(f"420x260+{cx}+{cy}")
        edit_dialog.deiconify()
        self._apply_title_bar_theme(edit_dialog)

        main_frame = tb.Frame(edit_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 名称输入
        tb.Label(main_frame, text="名称:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=0, column=0, padx=5, pady=8, sticky=tk.W
        )
        name_entry = tb.Entry(main_frame, width=25)
        name_entry.insert(0, name)
        name_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W + tk.E)

        # 主DNS输入
        tb.Label(main_frame, text="主DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=1, column=0, padx=5, pady=5, sticky=tk.W
        )
        primary_entry = tb.Entry(main_frame, width=25)
        primary_entry.insert(0, primary)
        primary_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W + tk.E)

        # 备用DNS输入
        tb.Label(main_frame, text="备用DNS:", font=(AppConfig.FONT_FAMILY, 10)).grid(
            row=2, column=0, padx=5, pady=5, sticky=tk.W
        )
        secondary_entry = tb.Entry(main_frame, width=25)
        secondary_entry.insert(0, secondary)
        secondary_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W + tk.E)

        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        def save_and_close():
            new_name = sanitize_input(name_entry.get().strip())
            new_primary = sanitize_input(primary_entry.get().strip())
            new_secondary = sanitize_input(secondary_entry.get().strip())

            if not new_name or not new_primary:
                messagebox.showerror("错误", "名称和主DNS不能为空")
                return

            on_save(new_name, new_primary, new_secondary)
            edit_dialog.destroy()

        button_frame = tb.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 5))

        tb.Button(button_frame, text="保存", command=save_and_close, bootstyle=SUCCESS).pack(
            side=tk.LEFT, padx=15
        )
        tb.Button(button_frame, text="取消", command=edit_dialog.destroy, bootstyle=SECONDARY).pack(
            side=tk.LEFT, padx=15
        )

        edit_dialog.lift()
        edit_dialog.attributes("-topmost", True)
        edit_dialog.focus_force()
        name_entry.focus()

    def show_info_dialog(self, title: str, message: str) -> None:
        """
        显示信息对话框

        Args:
            title: 对话框标题
            message: 信息消息
        """
        messagebox.showinfo(title, message)

    def show_warning_dialog(self, title: str, message: str) -> None:
        """
        显示警告对话框

        Args:
            title: 对话框标题
            message: 警告消息
        """
        messagebox.showwarning(title, message)

    def show_error_dialog(self, title: str, message: str) -> None:
        """
        显示错误对话框

        Args:
            title: 对话框标题
            message: 错误消息
        """
        messagebox.showerror(title, message)
