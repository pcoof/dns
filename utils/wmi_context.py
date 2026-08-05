"""WMI上下文管理器模块"""

import pythoncom
import win32com.client
from contextlib import contextmanager
from typing import Optional, Generator, Any


@contextmanager
def wmi_context(namespace: str = "root\\cimv2") -> Generator[Any, None, None]:
    """
    WMI连接的上下文管理器，自动处理初始化和清理

    Args:
        namespace: WMI命名空间，默认为 "root\\cimv2"

    Yields:
        WMI服务对象

    Raises:
        Exception: 当WMI连接失败时

    Example:
        >>> with wmi_context() as wmi:
        ...     adapters = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
        ...     for adapter in adapters:
        ...         print(adapter.Name)
    """
    pythoncom.CoInitialize()
    try:
        wmi = win32com.client.GetObject(f"winmgmts:\\\\.\\{namespace}")
        if not wmi:
            raise Exception(f"无法连接到WMI命名空间: {namespace}")
        yield wmi
    except Exception as e:
        raise Exception(f"WMI连接失败: {e}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@contextmanager
def wmi_dispatch_context(
    server: str = ".",
    namespace: str = "root\\cimv2",
) -> Generator[Any, None, None]:
    """
    使用Dispatch创建的WMI连接上下文管理器

    Args:
        server: 服务器名称，默认为本地 "."
        namespace: WMI命名空间，默认为 "root\\cimv2"

    Yields:
        WMI服务对象

    Raises:
        Exception: 当WMI连接失败时

    Example:
        >>> with wmi_dispatch_context() as wmi_service:
        ...     configs = wmi_service.ExecQuery(
        ...         "SELECT * FROM Win32_NetworkAdapterConfiguration"
        ...     )
    """
    pythoncom.CoInitialize()
    try:
        wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        wmi = wmi_service.ConnectServer(server, namespace)
        yield wmi
    except Exception as e:
        raise Exception(f"WMI Dispatch连接失败: {e}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
