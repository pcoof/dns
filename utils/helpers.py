"""Helper utility functions module"""

import ctypes
import os
import re
import socket
import sys
import tkinter as tk
from typing import Tuple, Optional


def center_window(window: tk.Tk, width: int, height: int) -> None:
    """
    Center the window on the screen

    Args:
        window: Tkinter window object
        width: Window width
        height: Window height

    Example:
        >>> root = tk.Tk()
        >>> center_window(root, 800, 600)
    """
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


def is_admin() -> bool:
    """
    Check if the program is running with administrator privileges

    Returns:
        bool: True if running as administrator, False otherwise

    Example:
        >>> if is_admin():
        ...     print("Running as administrator")
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_as_admin() -> bool:
    """
    Restart the program with administrator privileges

    Returns:
        bool: True if successfully started, False otherwise

    Example:
        >>> if not is_admin():
        ...     run_as_admin()
        ...     sys.exit(0)
    """
    try:
        script_path = os.path.abspath(sys.argv[0])
        if script_path.endswith(".py"):
            python_exe = sys.executable
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                python_exe,
                f'"{script_path}"',
                None,
                1,  # SW_SHOWNORMAL
            )
        else:
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                script_path,
                None,
                None,
                1,
            )
        if result > 32:
            print("Restarting with administrator privileges...")
            return True
        else:
            print(f"Failed to start with administrator privileges, error code: {result}")
            return False
    except Exception as e:
        print(f"Failed to start with administrator privileges: {e}")
        return False


def validate_dns_ip(ip: str) -> bool:
    """
    Strictly validate DNS IP address

    Args:
        ip: IP address string to validate

    Returns:
        bool: True if valid IPv4 or IPv6 address, False otherwise

    Example:
        >>> validate_dns_ip("8.8.8.8")
        True
        >>> validate_dns_ip("invalid")
        False
    """
    if not ip or not isinstance(ip, str):
        return False

    ip = ip.strip()

    # IPv4 validation
    ipv4_pattern = r"^(\d{1,3}\.){3}\d{1,3}$"
    if re.match(ipv4_pattern, ip):
        parts = ip.split(".")
        try:
            return all(0 <= int(part) <= 255 for part in parts)
        except ValueError:
            return False

    # IPv6 validation
    try:
        socket.inet_pton(socket.AF_INET6, ip)
        return True
    except (socket.error, AttributeError):
        return False


def sanitize_input(text: str) -> str:
    """
    Clean user input, remove potentially dangerous characters

    Args:
        text: Text to clean

    Returns:
        str: Cleaned text

    Example:
        >>> sanitize_input("test<script>alert('xss')</script>")
        'testscriptalert(xss)/script'
    """
    if not text or not isinstance(text, str):
        return ""
    # Remove HTML tags and special characters
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r'[<>"\'&]', "", text)
    return text.strip()


def is_ip(ip: str, version: Optional[int] = None) -> bool:
    """
    Check if it's a valid IP address

    Args:
        ip: IP address string to check
        version: Optional, specify IP version (4 or 6), default checks any valid IP

    Returns:
        bool: True if matches specified version, False otherwise

    Example:
        >>> is_ip("8.8.8.8", version=4)
        True
        >>> is_ip("2001:4860:4860::8888", version=6)
        True
    """
    if not ip:
        return version == 4  # Empty string is considered valid for IPv4

    def check_ipv4():
        try:
            parts = ip.split(".")
            return len(parts) == 4 and all(0 <= int(part) <= 255 for part in parts)
        except (ValueError, AttributeError):
            return False

    def check_ipv6():
        try:
            socket.inet_pton(socket.AF_INET6, ip)
            return True
        except (socket.error, AttributeError):
            return False

    if version == 4:
        return check_ipv4()
    elif version == 6:
        return check_ipv6()
    elif version is None:
        if not ip:
            return False
        return check_ipv4() or check_ipv6()
    else:
        return False


def get_notification_style(style: str) -> Tuple[str, str]:
    """
    Get notification style configuration

    Args:
        style: Style constant (SUCCESS, DANGER, WARNING, INFO)

    Returns:
        tuple: (prefix, color)

    Example:
        >>> from ttkbootstrap.constants import SUCCESS
        >>> prefix, color = get_notification_style(SUCCESS)
        >>> print(prefix, color)
        ✓  green
    """
    try:
        from ttkbootstrap.constants import INFO, SUCCESS, DANGER, WARNING
        style_map = {
            SUCCESS: ("✓ ", "green"),
            DANGER: ("✗ ", "red"),
            WARNING: ("⚠ ", "orange"),
            INFO: ("ℹ ", "blue"),
        }
    except ImportError:
        # If ttkbootstrap is not available, use default mapping
        style_map = {
            "SUCCESS": ("✓ ", "green"),
            "DANGER": ("✗ ", "red"),
            "WARNING": ("⚠ ", "orange"),
            "INFO": ("ℹ ", "blue"),
        }
    return style_map.get(style, ("ℹ ", "blue"))