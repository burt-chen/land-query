"""嵌入式包裝 — 讓 國土地籍批次查詢工具 跑在 Launcher 的分頁裡。

實作 create_frame(parent) -> ttk.Frame，由 Launcher 動態載入。
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk

_TOOL_ROOT = Path(__file__).parent


def _load_tool():
    """用 importlib 從絕對路徑載入 land_query.py，給唯一模組名避免衝突。"""
    spec = importlib.util.spec_from_file_location(
        "_land_query_tool", _TOOL_ROOT / "land_query.py"
    )
    mod = importlib.util.module_from_spec(spec)
    sys.modules["_land_query_tool"] = mod
    spec.loader.exec_module(mod)
    return mod


_tool = _load_tool()
_App = _tool.App


class _EmbeddedApp(_App):
    """把 App 嵌進任意 Tkinter widget（不需要 tk.Tk、不調整全域字型）。"""

    def __init__(self, parent: tk.Widget) -> None:
        # 直接呼叫父類，因為 App.__init__ 本來就只把 root 當 widget 用
        super().__init__(parent)


def create_frame(parent: tk.Widget) -> ttk.Frame:
    frame = ttk.Frame(parent)
    _EmbeddedApp(frame)
    return frame
