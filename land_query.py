#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""國土地籍批次查詢工具 (GUI 版)

從 LISP 下載地段代碼表，依 input.xlsx 的縣市/行政區/大段/小段/地號
逐筆查詢國土測繪中心，並輸出整理後的 Excel。

設定（網址、headless、逾時）放在同層的 config.json。
"""
from __future__ import annotations

import io
import json
import re
import sys
import threading
import time
import traceback
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText


# ===========================================================================
# 路徑 / 設定
# ===========================================================================

APP_TITLE = "國土地籍批次查詢工具"
UI_FONT_SIZE = 12

DEFAULT_CONFIG = {
    "section_url": "https://lisp.land.moi.gov.tw/MMS/Handle/DownloadQuerySection.ashx?DownloadType=xls",
    "landno_pattern": r"^\d{1,4}(-\d{1,4})?$",
    # 純 API 版（測試中）— 直接打 nlsc 的 API，不開瀏覽器
    "api_land_info_url": "https://api.nlsc.gov.tw/S09_Ralid/getLandInfoSect",
    "api_tile_index_url": "https://landmaps.nlsc.gov.tw/S_Maps/qryTileMapIndex",
    "api_location_query_url": "https://api.nlsc.gov.tw/MapSearch/LocationQuery",
    # 歷年國土利用調查（點查詢）。後面會接 /0/{經度}/{緯度}/4326，回 XML。
    # 這顆一定要帶 api_referer，不帶或亂帶會回 404 PERMISSION DENIED。
    "api_land_use_url": "https://api.nlsc.gov.tw/other/LandUsePointYears",
    "api_referer": "https://maps.nlsc.gov.tw/",
    "api_request_timeout": 20,
    "api_request_delay": 0.5,  # 每筆之間的禮貌延遲（太小易被 NLSC 限流，O 欄度分秒會空白）
}

INPUT_COLUMNS = ["縣市", "行政區", "大段", "小段", "地號"]

# 範例輸入資料（給「下載範例 input.xlsx」按鈕用，直接寫死）
SAMPLE_INPUT_ROWS = [
    ("高雄市", "橋頭區", "橋中段", "",     "92"),
    ("高雄市", "鳳山區", "埤頂段", "",     "2157-2"),
    ("高雄市", "苓雅區", "正文段", "",     "182"),
    ("高雄市", "林園區", "王公廟段", "",   "1013-1"),
    ("高雄市", "小港區", "港和段", "二小段", "446"),
    ("高雄市", "小港區", "港和段", "二小段", "446-1"),
    ("高雄市", "小港區", "港和段", "二小段", "447"),
    ("高雄市", "楠梓區", "藍田西段", "三小段", "2"),
]


def app_dir() -> Path:
    """程式所在資料夾。PyInstaller 打包後也能正確指向 .exe 旁邊。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    return app_dir() / "config.json"


def load_config() -> dict:
    """讀 config.json；缺檔或缺鍵時用預設值補齊。"""
    cfg = dict(DEFAULT_CONFIG)
    path = config_path()
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                cfg.update({k: v for k, v in data.items() if k in DEFAULT_CONFIG})
        except Exception:
            pass
    return cfg


def save_config(cfg: dict) -> None:
    config_path().write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def export_cols_path() -> Path:
    return app_dir() / "export_cols.json"


def _load_export_state() -> dict:
    """讀整份 export_cols.json，回 {active, active_preset_name, presets}。

    向後相容：舊版檔案是純 list，視為 active。
    """
    empty = {"active": None, "active_preset_name": "", "presets": {}}
    path = export_cols_path()
    if not path.exists():
        return dict(empty)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, list):  # 舊格式
            return {"active": data, "active_preset_name": "", "presets": {}}
        if not isinstance(data, dict):
            return dict(empty)
        active = data.get("active")
        presets = data.get("presets")
        name = data.get("active_preset_name")
        return {
            "active": active if isinstance(active, list) else None,
            "active_preset_name": name if isinstance(name, str) else "",
            "presets": presets if isinstance(presets, dict) else {},
        }
    except Exception:
        return dict(empty)


def _save_export_state(state: dict) -> bool:
    try:
        export_cols_path().write_text(
            json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


def _cols_to_serializable(cols: list[dict]) -> list[dict]:
    return [
        {
            "name": c["name"],
            "source": c.get("source"),
            "transform": c.get("transform"),
            "enabled": bool(c.get("enabled", True)),
        }
        for c in cols
    ]


def _normalize_export_cols(saved: list[dict]) -> list[dict]:
    """套模板對應：保留使用者順序/名稱/enabled，補新欄位、丟過時欄位。

    識別 key = (source, transform)，所以重命名不會影響對應。
    """
    def sig(c: dict) -> tuple:
        return (c.get("source"), c.get("transform"))

    template_by_sig = {sig(c): c for c in EXPORT_COLUMNS_TEMPLATE}
    seen: set[tuple] = set()
    result: list[dict] = []
    for c in saved:
        if not isinstance(c, dict):
            continue
        s = sig(c)
        tpl = template_by_sig.get(s)
        if tpl is None:  # 模板沒有 → 過時欄位
            continue
        result.append({
            "name": c.get("name") or tpl["name"],
            "source": tpl.get("source"),
            "transform": tpl.get("transform"),
            "enabled": bool(c.get("enabled", True)),
        })
        seen.add(s)
    # 模板有、檔案沒有的：附加到末尾，預設顯示
    for tpl in EXPORT_COLUMNS_TEMPLATE:
        if sig(tpl) not in seen:
            result.append({
                "name": tpl["name"],
                "source": tpl.get("source"),
                "transform": tpl.get("transform"),
                "enabled": True,
            })
    return result


def load_export_cols() -> list[dict]:
    """讀 active 欄位設定；缺檔或損毀就回預設。"""
    state = _load_export_state()
    saved = state.get("active")
    if not isinstance(saved, list):
        return [dict(c, enabled=True) for c in EXPORT_COLUMNS_TEMPLATE]
    return _normalize_export_cols(saved)


def save_export_cols(cols: list[dict], preset_name: str = "") -> bool:
    """寫入 active 欄位設定（保留現有 presets）+ 紀錄目前是哪個 preset 套上來的。"""
    state = _load_export_state()
    state["active"] = _cols_to_serializable(cols)
    state["active_preset_name"] = preset_name or ""
    return _save_export_state(state)


def load_active_preset_name() -> str:
    return _load_export_state().get("active_preset_name", "")


def list_export_presets() -> list[str]:
    return sorted(_load_export_state().get("presets", {}).keys())


def load_export_preset(name: str) -> list[dict] | None:
    presets = _load_export_state().get("presets", {})
    saved = presets.get(name)
    if not isinstance(saved, list):
        return None
    return _normalize_export_cols(saved)


def save_export_preset(name: str, cols: list[dict]) -> bool:
    if not name or not name.strip():
        return False
    state = _load_export_state()
    state.setdefault("presets", {})
    state["presets"][name] = _cols_to_serializable(cols)
    return _save_export_state(state)


def delete_export_preset(name: str) -> bool:
    state = _load_export_state()
    presets = state.get("presets", {})
    if name not in presets:
        return False
    del presets[name]
    return _save_export_state(state)


# ===========================================================================
# core：地段代碼下載 / 整理 / 比對 / 地號驗證（純邏輯，不依賴 GUI）
# ===========================================================================

@dataclass
class PreparedRow:
    """已對碼、可送 Selenium 查詢的一筆資料。"""
    city: str
    area: str
    section: str
    landno: str
    輸入縣市: str
    輸入行政區: str
    輸入大段: str
    輸入小段: str
    輸入地號: str
    # 地政事務所代碼（API 版用，如 "EF" = 岡山地政事務所）
    office: str = ""


@dataclass
class PreparedInput:
    valid: list[PreparedRow] = field(default_factory=list)
    no_code: list[dict] = field(default_factory=list)
    bad_landno: list[dict] = field(default_factory=list)


def _safe_str(v) -> str:
    """把 NaN、None 轉成空字串；其餘 str() 後 strip。"""
    if v is None:
        return ""
    # float NaN：NaN 不等於自己
    if isinstance(v, float) and v != v:
        return ""
    s = str(v).strip()
    if s in ("nan", "NaN", "NaT", "None"):
        return ""
    return s


def _read_xlsx_rows(source) -> tuple[list[str], list[dict]]:
    """讀 xlsx（第一個工作表），第一列當欄名。

    source 可以是檔案路徑或 BytesIO。
    回傳 (header, rows)；rows 是 list of dict，key 為欄名。
    空白列（整列都 None）會被跳過。
    """
    from openpyxl import load_workbook
    try:
        wb = load_workbook(source, read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(f"無法讀取 Excel（需 xlsx 格式）：{e}")
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        first = next(rows_iter)
    except StopIteration:
        wb.close()
        return [], []
    header = [(_safe_str(c) if c is not None else "") for c in first]
    out: list[dict] = []
    for row in rows_iter:
        if not row or all(c is None for c in row):
            continue
        d = {}
        for i, h in enumerate(header):
            if not h:
                continue
            d[h] = row[i] if i < len(row) else None
        out.append(d)
    wb.close()
    return header, out


def download_section_table(url: str) -> dict[str, dict]:
    """下載地段代碼表，整理成 key -> 段資料 的索引。

    每筆段資料含：縣市 / 行政區 / 事務所 / 大段 / 小段 / 所區碼 / 備註（給人看）、
    city / area / office / section（給查詢用的代碼）。
    """
    import urllib3
    import requests
    from openpyxl import load_workbook

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    resp = requests.get(url, verify=False, timeout=60)
    resp.raise_for_status()
    try:
        wb = load_workbook(io.BytesIO(resp.content), read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(
            f"LISP 下載回來的檔案不是 xlsx 格式，openpyxl 無法解析：{e}"
        )
    ws = wb.active

    index: dict[str, dict] = {}
    # 跳過第一列（標頭），用欄位位置取值，對應原本 raw.iloc[:, n]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(c is None for c in row):
            continue
        def _col(i: int) -> str:
            return _safe_str(row[i]) if i < len(row) else ""
        A = _col(0)  # 段
        B = _col(1)  # 小段
        C = _col(2)  # 代碼
        D = _col(3)  # 備註
        E = _col(4)  # 所區碼
        F = _col(5)  # 縣市名稱
        G = _col(6)  # 鄉鎮名稱
        H = _col(7)  # 事務所名稱

        key = F + G + A + B
        city = E[:1]
        # area = 縣市碼 + 鄉鎮市區代碼（如 E20）
        area = (E[:1] + E[-2:]) if len(E) >= 2 else city
        # office = 所區碼前 2 碼（如 EF）
        office = E[:2]
        try:
            section_num = int(float(C)) if C else 0
        except (ValueError, TypeError):
            section_num = 0
        section = str(section_num).zfill(4)

        index[key] = {
            "縣市": F, "行政區": G, "事務所": H, "大段": A, "小段": B,
            "所區碼": E, "備註": D,
            "city": city, "area": area, "office": office, "section": section,
        }
    wb.close()
    return index


def prepare_input(
    input_rows: list[dict],
    section_index: dict[str, dict],
    landno_pattern: str,
) -> PreparedInput:
    """讀使用者輸入（list of dict），套地段代碼，過濾找不到代碼 / 地號格式錯誤的資料。"""
    if input_rows:
        missing = [c for c in INPUT_COLUMNS if c not in input_rows[0]]
        if missing:
            raise ValueError(f"input.xlsx 缺少必要欄位: {', '.join(missing)}")

    pattern = re.compile(landno_pattern)
    result = PreparedInput()

    for row in input_rows:
        city_in = _safe_str(row.get("縣市"))
        area_in = _safe_str(row.get("行政區"))
        seg_in = _safe_str(row.get("大段"))
        sub_in = _safe_str(row.get("小段"))
        landno = _safe_str(row.get("地號"))

        # 對 key 用的：大段去尾 "段"、小段去尾 "小段"
        seg_norm = re.sub(r"段$", "", seg_in)
        sub_norm = re.sub(r"小段$", "", sub_in)
        key = city_in + area_in + seg_norm + sub_norm

        base_info = {
            "輸入縣市": city_in,
            "輸入行政區": area_in,
            "輸入大段": seg_in,
            "輸入小段": sub_in,
            "輸入地號": landno,
            "landno": landno,
        }

        sec = section_index.get(key)
        if not sec or not sec.get("city"):
            result.no_code.append(base_info)
            continue

        if not pattern.match(landno):
            result.bad_landno.append(base_info)
            continue

        result.valid.append(
            PreparedRow(
                city=sec["city"],
                area=sec["area"],
                section=sec["section"],  # 已經 zfill(4)
                landno=landno,
                輸入縣市=city_in,
                輸入行政區=area_in,
                輸入大段=seg_in,
                輸入小段=sub_in,
                輸入地號=landno,
                office=sec.get("office", ""),
            )
        )

    return result


# ===========================================================================
# 共用：停止信號（給 API 查詢用）
# ===========================================================================


class StopRequested(Exception):
    """使用者按下停止；查詢迴圈接到後安全結束。"""


# ===========================================================================
# core：純 API 查詢（不開瀏覽器，用 requests 直接打 nlsc 後端）
# ===========================================================================

# getLandInfoSect 回傳的 ralid 欄位代碼對照（資料來源：NLSC API）
# 不確定的代碼維持原 AAxx；確定有對應的就改成中文欄名。
RALID_FIELD_MAP = {
    "AA45": "縣市代碼",
    "AA46": "鄉鎮市區代碼",
    "AA48": "段代碼",
    "AA49": "地號代碼",
    "AA05": "登記日期",
    "AA06": "登記原因代碼",
    "AA08": "地目",
    "AA09": "等則",
    "AA10": "面積(平方公尺)",
    # AA11/AA12 在 ralid 是 base64 編碼，使用分區/使用地類別改從 land.AA11/AA12 直接取
    "AA16": "公告現值",
    "AA17": "公告地價",
    "AA21": "都市計畫面積",
    "AA22": "非都市計畫面積",
    "AA23": "使用分區",
    "AA24": "登記面積",
    "AA27": "公告日期",
}


# lcdetype.lcde_* 的代碼 → 中文類別 + 對應的 JSON key 清單
# （類別 1「本國人」由 m/f/o 三個子分類加總而成）
LCDETYPE_CATEGORIES = [
    ("1", "本國人",      ["lcde_1_m", "lcde_1_f", "lcde_1_o"]),
    ("2", "外國人",      ["lcde_2"]),
    ("3", "國有",        ["lcde_3"]),
    ("4", "省市有",      ["lcde_4"]),
    ("5", "縣市有",      ["lcde_5"]),
    ("6", "鄉鎮市有",    ["lcde_6"]),
    ("7", "本國私法人",  ["lcde_7"]),
    ("8", "外國法人",    ["lcde_8"]),
    ("9", "祭祀公業",    ["lcde_9"]),
    ("a", "其他",        ["lcde_a"]),
    ("b", "銀行法人",    ["lcde_b"]),
    ("c", "未知類別(c)", ["lcde_c"]),
    ("d", "大陸地區法人", ["lcde_d"]),
]


def _format_owner_type_breakdown(lcdetype: dict) -> str:
    """把 lcdetype 的所有權人類別比例串成單一字串。

    例如：「本國人:63.77%外國人:4.76%國有:31.47%」
    只列出比例 > 0 的類別；類別 1「本國人」由 m/f/o 三個子分類加總。
    """
    if not lcdetype:
        return ""
    parts: list[str] = []
    for _code, name, keys in LCDETYPE_CATEGORIES:
        total = 0.0
        for k in keys:
            try:
                total += float(lcdetype.get(k) or 0)
            except (TypeError, ValueError):
                pass
        if total > 0:
            parts.append(f"{name}:{total * 100:.2f}%")
    return "".join(parts)


# =============================================================================
# 「簡化版匯出」設定 — 對應「欄位調整範本.xlsx」工作表2 的 32 欄
# =============================================================================
# 每筆是一個 dict：
#   name      — Excel 最終欄名
#   source    — API 結果 dict 裡的 key；None 表示空白欄
#   transform — None / "minguo_date" / "yuan_per_sqm" / "frac_num" / "frac_den"
# =============================================================================

def _xform_minguo_date(v) -> str:
    """民國日期轉中文：'1011018' → '民國101年10月18日'"""
    s = str(v).strip()
    if not s or not s.isdigit() or len(s) not in (6, 7):
        return s
    # 6 碼: YYMMDD, 7 碼: YYYMMDD (民國 3 碼年)
    y = int(s[:-4])
    m = int(s[-4:-2])
    d = int(s[-2:])
    if not (1 <= m <= 12 and 1 <= d <= 31):
        return s
    return f"民國{y}年{m}月{d}日"


def _xform_yuan_per_sqm(v) -> str:
    """金額加單位：'7300' → '7300 元/平方公尺'"""
    s = str(v).strip()
    if not s or s in ("0", "0.0"):
        return ""
    return f"{s} 元/平方公尺"


def _xform_frac_num(v) -> str:
    """取分數的分子：'1/3' → '1'；不是分數就原樣回。"""
    s = str(v).strip()
    if "/" in s:
        return s.split("/", 1)[0].strip()
    return s


def _xform_frac_den(v) -> str:
    """取分數的分母：'1/3' → '3'；不是分數就回空。"""
    s = str(v).strip()
    if "/" in s:
        return s.split("/", 1)[1].strip()
    return ""


_TRANSFORMS = {
    "minguo_date": _xform_minguo_date,
    "yuan_per_sqm": _xform_yuan_per_sqm,
    "frac_num": _xform_frac_num,
    "frac_den": _xform_frac_den,
}


# (source, transform) → (來源描述, 處理描述) for「欄位設定」對話框
# 跟 api_field_mapping.md 的 30 欄總表保持一致；新增/改 EXPORT_COLUMNS_TEMPLATE 時記得同步
COLUMN_DESCRIPTIONS: dict[tuple, tuple[str, str]] = {
    ("輸入縣市", None):                ("使用者 input.xlsx", ""),
    ("輸入行政區", None):              ("使用者 input.xlsx", ""),
    ("輸入大段", None):                ("使用者 input.xlsx", ""),
    ("輸入小段", None):                ("使用者 input.xlsx", ""),
    ("輸入地號", None):                ("使用者 input.xlsx", ""),
    ("面積(平方公尺)", None):          ("ralid.AA10", ""),
    ("使用分區", None):                ("land.AA11", ""),
    ("使用地類別", None):              ("land.AA12", ""),
    ("登記日期", "minguo_date"):       ("ralid.AA05", "轉民國年月日"),
    ("公告現值", "yuan_per_sqm"):      ("ralid.AA16", "加單位「元/平方公尺」"),
    ("公告地價", "yuan_per_sqm"):      ("ralid.AA17", "加單位「元/平方公尺」"),
    ("權利人類別", None):              ("lcdetype.lcde_*", "組字串「本國人:63.77%外國人:4.76%國有:31.47%」"),
    ("地籍連結(JSONP)", None):         ("qryTileMapIndex cx,cy", "組 http://maps.nlsc.gov.tw/go/{cy}/{cx}"),
    ("行政區", None):                  ("LocationQuery 解析", ""),
    ("經緯度(JSONP)", None):           ("qryTileMapIndex cx,cy", "直接組 cx,cy"),
    ("經緯度(度分秒)", None):          ("LocationQuery 解析", ""),
    ("國土利用_年月", None):           ("LandUsePointYears 最新一期", "YEAR 年 LMONTH 月"),
    ("國土利用_現況", None):           ("LandUsePointYears 最新一期", "LCODE-NAME"),
    ("TWD97_E", None):                 ("從 cx,cy 純 Python 換算", "E 座標(公尺,四捨五入到整數)"),
    ("TWD97_N", None):                 ("從 cx,cy 純 Python 換算", "N 座標(公尺,四捨五入到整數)"),
    ("TWD97", None):                   ("從 cx,cy 純 Python 換算", "格式 E:xxx N:xxx,EPSG:3826"),
    ("地號(JSONP組合)", None):         ("組合字串", "{所}所({office}{sect}){段}{地號}地號"),
    ("所有人_姓名", None):             ("land.userList[0].name", ""),
    ("所有人_身分證號", None):         ("land.userList[0].id", ""),
    ("所有人_類型", None):             ("land.userList[0].type", ""),
    ("所有人_範圍", None):             ("land.userList[0].scope", ""),
    ("所有人_持分", "frac_den"):       ("land.userList[0].denominator", "從 1/3 取分母"),
    ("所有人_持分", "frac_num"):       ("land.userList[0].numerator", "從 1/3 取分子"),
    ("所有人_公告現值", "yuan_per_sqm"): ("land.userList[0].price", "加單位「元/平方公尺」"),
    ("所有人_管理機關", None):         ("land.userList[0].manage", ""),
    ("查詢縣市", None):                ("LISP 段碼表對碼", ""),
    ("查詢區", None):                  ("LISP 段碼表對碼", ""),
    ("查詢地段", None):                ("LISP 段碼表對碼", ""),
    ("查詢地號", None):                ("LISP 段碼表對碼", ""),
}


EXPORT_COLUMNS_TEMPLATE = [
    {"name": "輸入縣市",          "source": "輸入縣市"},
    {"name": "輸入行政區",        "source": "輸入行政區"},
    {"name": "輸入大段",          "source": "輸入大段"},
    {"name": "輸入小段",          "source": "輸入小段"},
    {"name": "輸入地號",          "source": "輸入地號"},
    {"name": "面積",              "source": "面積(平方公尺)"},
    {"name": "使用分區",          "source": "使用分區"},
    {"name": "使用地類別",        "source": "使用地類別"},
    {"name": "登記日期",          "source": "登記日期",       "transform": "minguo_date"},
    {"name": "公告現值",          "source": "公告現值",       "transform": "yuan_per_sqm"},
    {"name": "公告地價",          "source": "公告地價",       "transform": "yuan_per_sqm"},
    {"name": "權利人類別",        "source": "權利人類別"},
    {"name": "地籍連結",          "source": "地籍連結(JSONP)"},
    {"name": "行政區",            "source": "行政區"},
    {"name": "經緯度(度)",        "source": "經緯度(JSONP)"},
    {"name": "經緯度(度分秒)",    "source": "經緯度(度分秒)"},
    {"name": "國土利用_年月",     "source": "國土利用_年月"},
    {"name": "國土利用_現況",     "source": "國土利用_現況"},
    {"name": "TWD97(E)",          "source": "TWD97_E"},
    {"name": "TWD97(N)",          "source": "TWD97_N"},
    {"name": "TWD97",             "source": "TWD97"},  # 純 Python 從 cx,cy 換算（E:xxx N:xxx）
    {"name": "地號",              "source": "地號(JSONP組合)"},
    {"name": "所有權人",          "source": "所有人_姓名"},
    {"name": "統一編號",          "source": "所有人_身分證號"},
    {"name": "所有權人類別",      "source": "所有人_類型"},
    {"name": "權利範圍類別",      "source": "所有人_範圍"},
    {"name": "權利範圍持分_分母", "source": "所有人_持分",    "transform": "frac_den"},
    {"name": "權利範圍持分_分子", "source": "所有人_持分",    "transform": "frac_num"},
    {"name": "申報地價",          "source": "所有人_公告現值", "transform": "yuan_per_sqm"},
    {"name": "管理者名稱",        "source": "所有人_管理機關"},
    {"name": "查詢縣市",          "source": "查詢縣市"},
    {"name": "查詢區",            "source": "查詢區"},
    {"name": "查詢地段",          "source": "查詢地段"},
    {"name": "查詢地號",          "source": "查詢地號"},
]


def export_results_template(
    results: list[dict],
    path: str,
    columns: list[dict] | None = None,
) -> None:
    """依 EXPORT_COLUMNS_TEMPLATE 的欄位設定 + 資料處理 匯出 Excel。

    來源欄不存在或值為 None/'' 時填空白；transform 套用後也可能是空白。
    """
    from openpyxl import Workbook
    cols = list(columns) if columns is not None else list(EXPORT_COLUMNS_TEMPLATE)

    wb = Workbook()
    ws = wb.active
    ws.append([c["name"] for c in cols])
    for r in results:
        out_row = []
        for spec in cols:
            src = spec.get("source")
            tx = spec.get("transform")
            if src is None:
                out_row.append("")
                continue
            val = r.get(src, "")
            if val is None:
                val = ""
            if tx and val != "":
                fn = _TRANSFORMS.get(tx)
                if fn is not None:
                    try:
                        val = fn(val)
                    except Exception:
                        pass
            out_row.append(val)
        ws.append(out_row)
    wb.save(path)
    wb.close()


def _wgs84_to_twd97(lon: float, lat: float) -> tuple[float, float] | None:
    """WGS84 (EPSG:4326) → TWD97 二度分帶 (EPSG:3826)。

    純 Python 實作（不依賴 pyproj），用 GRS80 橢球 + 121°E 中央子午線。
    精度約 ±1 公尺，對顯示用途夠用；想要 mm 級精度才換 pyproj。
    """
    from math import radians, sin, cos, tan, sqrt
    try:
        # GRS80 ellipsoid
        a = 6378137.0
        f = 1.0 / 298.257222101
        e2 = 2 * f - f * f
        # TWD97 二度分帶 EPSG:3826 參數
        k0 = 0.9999
        lon0 = radians(121.0)   # 中央子午線
        x0 = 250000.0           # false easting
        y0 = 0.0                # false northing

        phi = radians(lat)
        lam = radians(lon)

        ep2 = e2 / (1 - e2)
        N = a / sqrt(1 - e2 * sin(phi) ** 2)
        T = tan(phi) ** 2
        C = ep2 * cos(phi) ** 2
        A = (lam - lon0) * cos(phi)

        M = a * (
            (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256) * phi
            - (3 * e2 / 8 + 3 * e2 ** 2 / 32 + 45 * e2 ** 3 / 1024) * sin(2 * phi)
            + (15 * e2 ** 2 / 256 + 45 * e2 ** 3 / 1024) * sin(4 * phi)
            - (35 * e2 ** 3 / 3072) * sin(6 * phi)
        )

        x = k0 * N * (
            A + (1 - T + C) * A ** 3 / 6
            + (5 - 18 * T + T ** 2 + 72 * C - 58 * ep2) * A ** 5 / 120
        ) + x0
        y = k0 * (
            M + N * tan(phi) * (
                A ** 2 / 2
                + (5 - T + 9 * C + 4 * C ** 2) * A ** 4 / 24
                + (61 - 58 * T + T ** 2 + 600 * C - 330 * ep2) * A ** 6 / 720
            )
        ) + y0
        return x, y
    except Exception:
        return None


def _parse_location_query(text: str) -> dict:
    """解析 LocationQuery 回的 '<br>' 分隔字串。

    範例輸入：
      E@行政區:高雄市橋頭區橋頭里<br>經緯度:120.310188,22.756207   (度)<br>...

    回傳：{'行政區': ..., '經緯度(度)': ..., '經緯度(度分秒)': ..., '國土利用現況': ...}
    """
    result = {}
    # 去掉開頭的 'E@' 前綴
    s = text.strip()
    if s.startswith("E@"):
        s = s[2:]
    parts = [p.strip() for p in s.split("<br>") if p.strip()]
    for p in parts:
        if ":" not in p:
            continue
        key, val = p.split(":", 1)
        key, val = key.strip(), val.strip()
        if key == "行政區":
            result["行政區"] = val
        elif key == "經緯度":
            if "(度分秒)" in val:
                result["經緯度(度分秒)"] = val.replace("(度分秒)", "").strip()
            elif "(度)" in val:
                result["經緯度(度)"] = val.replace("(度)", "").strip()
            else:
                # 沒帶 (度) 就當度
                result.setdefault("經緯度(度)", val)
        elif "國土利用現況" in key:
            result["國土利用現況"] = val
    return result


def _parse_land_use_years(xml_text: str) -> dict:
    """解析 LandUsePointYears 回的 XML，取**最新一期**國土利用調查。

    回應長這樣（民國年由小到大，但不保證，所以用 YEAR 取最大的那筆）：

      <root>
        <ITEM><YEAR>112</YEAR><LYEAR>2023</LYEAR><LMONTH>10</LMONTH>
              <LCODE>090501</LCODE><NAME>未使用地</NAME>...</ITEM>
        <ITEM><YEAR>114</YEAR><LYEAR>2025</LYEAR><LMONTH>7</LMONTH>...</ITEM>
      </root>

    查無資料時是 <root><CONTENT>無任何資料</CONTENT></root>。

    回傳：{'國土利用_年月': '114年7月', '國土利用_現況': '090501-未使用地'}
    缺欄位就給空字串；整份解析失敗回 {}。

    ⚠️ 不能用 LYEAR 排序：NLSC 資料有瑕疵（民國 95 那筆的 LYEAR 標成 2023），
       只有 YEAR（民國年）是可信的。
    """
    if not xml_text or not xml_text.strip():
        return {}
    try:
        root = ET.fromstring(xml_text)
    except Exception:
        return {}

    latest = None
    latest_year = None
    for item in root.findall("ITEM"):
        raw_year = (item.findtext("YEAR") or "").strip()
        try:
            year = int(raw_year)
        except ValueError:
            continue
        if latest_year is None or year > latest_year:
            latest_year, latest = year, item
    if latest is None:
        return {}

    month = (latest.findtext("LMONTH") or "").strip()
    lcode = (latest.findtext("LCODE") or "").strip()
    name = (latest.findtext("NAME") or "").strip()

    # 年月：有月份才加「月」（民國 82 那期沒有 LMONTH）
    ym = f"{latest_year}年{month}月" if month else f"{latest_year}年"
    # 現況：LCODE-NAME；只有一邊有值就給那一邊
    if lcode and name:
        status = f"{lcode}-{name}"
    else:
        status = lcode or name

    return {"國土利用_年月": ym, "國土利用_現況": status}


def _api_format_land_record(
    row: "PreparedRow",
    payload: dict,
    location_text: str | None = None,
    tile_index: dict | None = None,
    land_use_xml: str | None = None,
) -> dict:
    """把所有 API 回的資料攤平成跟 Selenium 版相容的 dict。"""
    data = {
        "輸入縣市": row.輸入縣市, "輸入行政區": row.輸入行政區,
        "輸入大段": row.輸入大段, "輸入小段": row.輸入小段, "輸入地號": row.輸入地號,
        "查詢縣市": row.city, "查詢區": row.area,
        "查詢地段": row.section, "查詢地號": row.landno,
    }

    ralid = payload.get("ralid") or {}
    for k, v in ralid.items():
        # AA11/AA12 在 ralid 是 base64 編碼，跳過；改從 land.AA11/AA12 取純文字（見下方）
        if k in ("AA11", "AA12"):
            continue
        col = RALID_FIELD_MAP.get(k, k)
        data[col] = v

    # land 區塊的 AA11/AA12 是中文版「使用分區 / 使用地類別」
    # （ralid 區塊的 AA11/AA12 是同樣兩個欄位的 base64 編碼，不用）
    land_block = payload.get("land") or {}
    data["使用分區"] = land_block.get("AA11", "")
    data["使用地類別"] = land_block.get("AA12", "")

    # lcdetype 其實是「所有權人類別面積比例」，不是土地使用分類
    lcdetype = payload.get("lcdetype") or {}
    for k, v in lcdetype.items():
        data[f"lcdetype.{k}"] = v
    # 組成 Excel「權利人類別」欄用的字串
    data["權利人類別"] = _format_owner_type_breakdown(lcdetype)

    # 所有人清單（取代「公有土地」）
    user_list = ((payload.get("land") or {}).get("userList")) or []
    if user_list:
        for i, u in enumerate(user_list, start=1):
            prefix = f"所有人{i}" if len(user_list) > 1 else "所有人"
            data[f"{prefix}_姓名"] = u.get("name", "")
            data[f"{prefix}_身分證號"] = u.get("id", "")
            data[f"{prefix}_類型"] = u.get("type", "")
            data[f"{prefix}_持分"] = (
                f"{u.get('numerator', '')}/{u.get('denominator', '')}"
                if u.get("denominator") else u.get("scope", ""))
            data[f"{prefix}_範圍"] = u.get("scope", "")
            data[f"{prefix}_公告現值"] = u.get("price", "")
            data[f"{prefix}_管理機關"] = u.get("manage", "")
        # 公有土地：type 不是「私有」/「未登錄」就算（國有、省有、市有、縣有、鄉鎮市有都算）
        public_owners = [
            u for u in user_list
            if str(u.get("type", "")).strip() not in ("", "私有", "未登錄")
        ]
        data["是否含公有土地"] = "是" if public_owners else "否"
        data["公有土地筆數"] = len(public_owners)
    else:
        data["是否含公有土地"] = "否"
        data["公有土地筆數"] = 0

    # 建物清單
    build_list = payload.get("buildList") or []
    data["建物筆數"] = len(build_list)

    # LocationQuery 解析（行政區、經緯度、國土利用）— 蓋過上面用輸入欄拼的行政區
    if location_text:
        for k, v in _parse_location_query(location_text).items():
            data[k] = v

    # 行政區若 LocationQuery 沒回，退而求其次用輸入欄拼
    if not data.get("行政區") and (row.輸入縣市 or row.輸入行政區):
        data["行政區"] = f"{row.輸入縣市}{row.輸入行政區}"

    # 經緯度(度) 若 LocationQuery 沒給就從 tile_index 的 cx,cy 直接組
    if tile_index and "cx" in tile_index and "cy" in tile_index:
        if not data.get("經緯度(度)"):
            data["經緯度(度)"] = f"{tile_index['cx']},{tile_index['cy']}"
        # TWD97 從 cx,cy 換算（純 Python，不需 pyproj）
        twd97 = _wgs84_to_twd97(tile_index["cx"], tile_index["cy"])
        if twd97:
            e, n = twd97
            data["TWD97_E"] = int(round(e))
            data["TWD97_N"] = int(round(n))
            data["TWD97"] = f"E:{int(round(e))} N:{int(round(n))}"

    # 從 tile_index 組出 3 個衍生字串欄位（給 EXPORT_COLUMNS_TEMPLATE 用）
    if tile_index:
        import base64
        cx = tile_index.get("cx")
        cy = tile_index.get("cy")
        office = tile_index.get("office", "") or ""
        sect = tile_index.get("sect", "") or ""
        office_str_b64 = tile_index.get("officeStr", "") or ""
        sect_str_b64 = tile_index.get("sectStr", "") or ""

        # 經緯度(JSONP) — 跟「經緯度(度)」可能相同，獨立欄位方便切換來源
        if cx is not None and cy is not None:
            data["經緯度(JSONP)"] = f"{cx},{cy}"
            # 地籍連結（NLSC 的 go/ 連結，會在地圖上定位到該點）
            data["地籍連結(JSONP)"] = f"http://maps.nlsc.gov.tw/go/{cy}/{cx}"

        # 地號完整字串：「岡山所(EF2424)橋中段92地號」
        try:
            office_str = base64.b64decode(office_str_b64).decode("utf-8") if office_str_b64 else ""
            sect_str = base64.b64decode(sect_str_b64).decode("utf-8") if sect_str_b64 else ""
            if office_str and sect_str:
                data["地號(JSONP組合)"] = (
                    f"{office_str}所({office}{sect}){sect_str}{row.輸入地號}地號"
                )
        except Exception:
            pass

    # 歷年國土利用調查（LandUsePointYears）— 只取最新一期的年月 + 現況
    data["國土利用_年月"] = ""
    data["國土利用_現況"] = ""
    if land_use_xml:
        for k, v in _parse_land_use_years(land_use_xml).items():
            data[k] = v

    return data


def run_api_query(
    rows: list[PreparedRow],
    cfg: dict,
    log: Callable[[str], None],
    progress: Callable[[int, int], None],
    should_stop: Callable[[], bool],
    on_row: Callable[[dict], None] | None = None,
    enable_land_use: bool = False,
) -> list[dict]:
    """純 API 查詢，每筆最多打 4 顆 API：
      1. getLandInfoSect — 土地基本資訊 + 所有人 + 公有土地
      2. qryTileMapIndex — 地塊中心經緯度（給 3、4 用）
      3. LocationQuery — 行政區 + 經緯度(度/度分秒) + 國土利用
      4. LandUsePointYears — 歷年國土利用調查（取最新一期的年月 + 現況）

    第 4 顆預設**不打**（`enable_land_use=False`），「國土利用_年月 / _現況」兩欄會空白；
    要那兩欄才傳 `enable_land_use=True`，每筆約多 0.3 秒（整批約慢一倍）。
    """
    import requests
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    session = requests.Session()
    session.headers.update({
        "Referer": cfg.get("api_referer", DEFAULT_CONFIG["api_referer"]),
        "Origin": "https://maps.nlsc.gov.tw",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "X-Requested-With": "XMLHttpRequest",
    })

    timeout = float(cfg.get("api_request_timeout", 20))
    delay = float(cfg.get("api_request_delay", 0.1))
    land_url = cfg.get("api_land_info_url", DEFAULT_CONFIG["api_land_info_url"])
    tile_url = cfg.get("api_tile_index_url", DEFAULT_CONFIG["api_tile_index_url"])
    loc_url = cfg.get("api_location_query_url", DEFAULT_CONFIG["api_location_query_url"])
    landuse_url = cfg.get("api_land_use_url", DEFAULT_CONFIG["api_land_use_url"])

    def get_tile_index(office: str, sect: str, landno8: str,
                       trace: list | None = None) -> dict | None:
        params = {
            "type": "2", "flag": "2",
            "office": office, "sect": sect, "landno": landno8,
            "alpah": "0.5f",
        }
        entry = {"seq": 2, "name": "qryTileMapIndex", "method": "GET",
                 "url": tile_url, "params": dict(params), "raw_key": "_raw_tile"}
        t0 = time.time()
        try:
            r = session.get(tile_url, params=params, timeout=timeout, verify=False)
            entry["elapsed"] = time.time() - t0
            entry["status"] = r.status_code
            entry["bytes"] = len(r.content)
            entry["url"] = r.url
            if r.status_code != 200:
                entry["error"] = f"HTTP {r.status_code}"
                return None
            arr = r.json()
            if isinstance(arr, list) and arr:
                return arr[0]
            if isinstance(arr, dict):
                return arr
            entry["error"] = "回應不是預期的 list/dict"
        except Exception as e:
            entry.setdefault("elapsed", time.time() - t0)
            entry["error"] = f"{type(e).__name__}: {e}"
            log(f"  qryTileMapIndex 失敗: {type(e).__name__}: {e}")
        finally:
            if trace is not None:
                trace.append(entry)
        return None

    def get_location_query(cx: float, cy: float, max_retry: int = 2,
                           trace: list | None = None) -> str | None:
        # 注意：LocationQuery 不能跟主 session 共用！
        # NLSC 後端怪規則：同一個 HTTP session 只回第一次完整資料，之後一律空白。
        # 改用一次性 request；空字串回應視為失敗、重試最多 max_retry 次。
        headers = {
            "Referer": cfg.get("api_referer", DEFAULT_CONFIG["api_referer"]),
            "Origin": "https://maps.nlsc.gov.tw",
            "User-Agent": session.headers.get("User-Agent", "Mozilla/5.0"),
            "X-Requested-With": "XMLHttpRequest",
        }
        params = {"center": f"{cx},{cy}"}
        entry = {"seq": 3, "name": "LocationQuery", "method": "POST",
                 "url": loc_url, "params": dict(params),
                 "raw_key": "_raw_location"}
        t0 = time.time()
        try:
            for attempt in range(max_retry + 1):
                try:
                    r = requests.post(
                        loc_url, data=params,
                        headers=headers, timeout=timeout, verify=False,
                    )
                    r.encoding = "utf-8"
                    entry["status"] = r.status_code
                    entry["bytes"] = len(r.content)
                    if r.status_code == 200 and r.text.strip():
                        if attempt:
                            entry["note"] = f"第 {attempt + 1} 次嘗試才成功（前 {attempt} 次空回應）"
                        return r.text
                    # 空字串：等一下再試
                    if attempt < max_retry:
                        time.sleep(0.5 + attempt * 0.5)
                    else:
                        entry["error"] = "空回應（重試 %d 次都沒拿到資料）" % max_retry
                except Exception as e:
                    if attempt == max_retry:
                        entry["error"] = f"{type(e).__name__}: {e}"
                        log(f"  LocationQuery 失敗 ({type(e).__name__}): {e}")
                        return None
                    time.sleep(0.5 + attempt * 0.5)
            return None
        finally:
            entry["elapsed"] = time.time() - t0
            entry["attempts"] = attempt + 1
            if trace is not None:
                trace.append(entry)

    def get_land_use_years(cx: float, cy: float,
                           trace: list | None = None) -> str | None:
        """歷年國土利用調查（點查詢），回 XML 字串。

        路徑格式：{base}/0/{經度}/{緯度}/4326
        ⚠️ 這顆一定要帶 Referer（session 已經帶了）；不帶會回 404 PERMISSION DENIED。
        """
        url = f"{landuse_url.rstrip('/')}/0/{cx}/{cy}/4326"
        entry = {"seq": 4, "name": "LandUsePointYears", "method": "GET",
                 "url": url, "params": None, "raw_key": "_raw_land_use",
                 "note": "參數在路徑上：/0/{經度}/{緯度}/4326（srid 4326 = WGS84）"}
        t0 = time.time()
        try:
            r = session.get(url, timeout=timeout, verify=False)
            entry["elapsed"] = time.time() - t0
            entry["status"] = r.status_code
            entry["bytes"] = len(r.content)
            if r.status_code != 200:
                entry["error"] = f"HTTP {r.status_code}（可能是 Referer 被擋）"
                log(f"  LandUsePointYears HTTP {r.status_code}（可能是 Referer 被擋）")
                return None
            r.encoding = "utf-8"
            return r.text
        except Exception as e:
            entry.setdefault("elapsed", time.time() - t0)
            entry["error"] = f"{type(e).__name__}: {e}"
            log(f"  LandUsePointYears 失敗: {type(e).__name__}: {e}")
        finally:
            if trace is not None:
                trace.append(entry)
        return None

    all_results: list[dict] = []
    total = len(rows)

    for idx, row in enumerate(rows):
        if should_stop():
            raise StopRequested()
        progress(idx, total)

        # 把 input 的地號（如「45-1」或「123」）轉成 API 的 8 碼格式
        # 「123」 → 「01230000」、「45-1」 → 「00450001」
        try:
            if "-" in row.landno:
                main, sub = row.landno.split("-", 1)
            else:
                main, sub = row.landno, "0"
            landno8 = f"{int(main):04d}{int(sub):04d}"
        except Exception:
            log(f"  第 {idx + 1} 筆地號格式異常: {row.landno}")
            data = {
                "輸入縣市": row.輸入縣市, "輸入行政區": row.輸入行政區,
                "輸入大段": row.輸入大段, "輸入小段": row.輸入小段, "輸入地號": row.輸入地號,
                "查詢狀態": f"地號格式異常: {row.landno}",
                "_api_trace": [{
                    "seq": 0, "name": "（沒打任何 API）", "skipped": True,
                    "note": f"地號「{row.landno}」轉不成 8 碼格式，在送出前就擋下來了",
                }],
            }
            all_results.append(data)
            if on_row:
                try: on_row(data)
                except Exception: pass
            continue

        log(f"查詢第 {idx + 1}/{total} 筆: {row.輸入縣市} {row.輸入行政區} {row.輸入大段}{row.輸入小段} {row.landno}  (API)")

        # --- 1. 主資料 ---
        # trace：這一筆打了哪幾顆 API、送什麼參數、回什麼，給右鍵「API 原始回應」用
        trace: list[dict] = []
        land_params = {"city": row.city, "sect": row.section, "landno": landno8}
        land_entry = {"seq": 1, "name": "getLandInfoSect", "method": "POST",
                      "url": land_url, "params": dict(land_params),
                      "raw_key": "_raw_payload"}
        trace.append(land_entry)
        t0 = time.time()
        try:
            r = session.post(
                land_url, data=land_params, timeout=timeout, verify=False,
            )
            land_entry["elapsed"] = time.time() - t0
            land_entry["status"] = r.status_code
            land_entry["bytes"] = len(r.content)
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}")
            payload = r.json()
        except Exception as e:
            land_entry.setdefault("elapsed", time.time() - t0)
            land_entry["error"] = f"{type(e).__name__}: {e}"
            for seq, name in ((2, "qryTileMapIndex"), (3, "LocationQuery"),
                              (4, "LandUsePointYears")):
                trace.append({"seq": seq, "name": name, "skipped": True,
                              "note": "跳過：getLandInfoSect 失敗，這筆不再往下打"})
            data = {
                "輸入縣市": row.輸入縣市, "輸入行政區": row.輸入行政區,
                "輸入大段": row.輸入大段, "輸入小段": row.輸入小段, "輸入地號": row.輸入地號,
                "查詢縣市": row.city, "查詢區": row.area,
                "查詢地段": row.section, "查詢地號": row.landno,
                "查詢狀態": f"getLandInfoSect 失敗: {type(e).__name__}: {e}",
                "_api_trace": trace,
            }
            all_results.append(data)
            if on_row:
                try: on_row(data)
                except Exception: pass
            time.sleep(delay)
            continue

        # --- 2. tile index 拿地塊中心 ---
        # office = row.office（地政事務所代碼，如 'EF' = 岡山）
        # 舊資料 office 可能為空，fallback 用 row.area（雖然會失敗，至少不會 crash）
        office_code = row.office or row.area
        tile_index = get_tile_index(office_code, row.section, landno8, trace=trace)

        # --- 3. LocationQuery 拿行政區/經緯度 ---
        location_text = None
        land_use_xml = None
        if tile_index and "cx" in tile_index and "cy" in tile_index:
            location_text = get_location_query(
                tile_index["cx"], tile_index["cy"], trace=trace)
            # --- 4. LandUsePointYears 拿歷年國土利用（取最新一期）---
            if enable_land_use:
                land_use_xml = get_land_use_years(
                    tile_index["cx"], tile_index["cy"], trace=trace)
            else:
                trace.append({
                    "seq": 4, "name": "LandUsePointYears", "skipped": True,
                    "note": "跳過：執行分頁沒勾「歷年國土利用」"
                            "（「國土利用_年月 / _現況」兩欄會空白）",
                })
        else:
            # 沒拿到地塊中心座標，3、4 顆沒東西可餵，直接跳過
            why = ("跳過：qryTileMapIndex 沒回 cx,cy（第 3、4 顆要用這組座標當參數）")
            for seq, name in ((3, "LocationQuery"), (4, "LandUsePointYears")):
                trace.append({"seq": seq, "name": name, "skipped": True, "note": why})

        # --- 攤平 ---
        data = _api_format_land_record(
            row, payload, location_text, tile_index, land_use_xml)

        # 保留 4 顆 API 的原始回應，給 GUI 右鍵「檢視 API 回應」用
        # 用底線開頭，export_results_template 不會誤抓到（它只看 EXPORT_COLUMNS_TEMPLATE）
        data["_raw_payload"] = payload
        data["_raw_tile"] = tile_index
        data["_raw_location"] = location_text
        data["_raw_land_use"] = land_use_xml
        data["_api_trace"] = trace

        if not (payload.get("ralid") or payload.get("land", {}).get("userList")):
            data["查詢狀態"] = "查無資料"
        else:
            data.setdefault("查詢狀態", "成功")

        all_results.append(data)
        if on_row:
            try: on_row(data)
            except Exception: pass

        time.sleep(delay)

    progress(total, total)
    return all_results


# ===========================================================================
# GUI：分頁版主視窗
# ===========================================================================

def _configure_global_fonts(size: int = UI_FONT_SIZE) -> None:
    """獨立執行時調大字型；嵌入模式由 launcher 控制，這個不會被呼叫。"""
    import tkinter.font as tkfont
    for name in (
        "TkDefaultFont", "TkTextFont", "TkFixedFont", "TkMenuFont",
        "TkHeadingFont", "TkCaptionFont", "TkSmallCaptionFont",
        "TkIconFont", "TkTooltipFont",
    ):
        try:
            tkfont.nametofont(name).configure(size=size)
        except tk.TclError:
            pass
    style = ttk.Style()
    for st in (
        "TButton", "TLabel", "TEntry", "TCombobox", "TCheckbutton",
        "TRadiobutton", "TMenubutton", "TNotebook", "TNotebook.Tab",
        "TLabelframe", "TLabelframe.Label", "Treeview", "Treeview.Heading",
        "TProgressbar",
    ):
        try:
            style.configure(st, font=("TkDefaultFont", size))
        except tk.TclError:
            pass
    try:
        style.configure("Treeview", rowheight=int(size * 2.0))
    except tk.TclError:
        pass


class App:
    """主視窗（核心 UI 邏輯）。

    self.root 可能是 tk.Tk 或任意 widget（嵌入模式）。
    """

    def __init__(self, root: tk.Widget) -> None:
        self.root = root
        self.cfg = load_config()

        self._input_path = tk.StringVar(value="")
        self._output_path = tk.StringVar(value="")

        self._prepared: PreparedInput | None = None
        self._section_index: dict[str, dict] = {}  # 地段代碼表（給「段名代碼表」分頁）
        # 三個下拉的完整選項清單（打字過濾時的母清單）
        self._sect_city_master: list[str] = []
        self._sect_dist_master: list[str] = []
        self._sect_seg_master: list[str] = []
        self._running = False
        self._stop_flag = False
        self._worker: threading.Thread | None = None
        # API 查詢狀態
        self._running_api = False
        self._stop_flag_api = False
        self._worker_api: threading.Thread | None = None
        self._results_api_done: list[dict] = []
        self._results_api_fail: list[dict] = []
        # 可調整的匯出欄位設定（顯示/順序/欄位名），優先讀 export_cols.json
        self._export_cols: list[dict] = load_export_cols()

        self._build_ui()

    # ---- UI 建構 ----

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.notebook.add(self._build_files_tab(self.notebook), text="1. 檔案與預覽")
        self.notebook.add(self._build_api_run_tab(self.notebook), text="2. 執行")
        self.notebook.add(self._build_log_tab(self.notebook), text="3. 日誌")
        self.notebook.add(self._build_settings_tab(self.notebook), text="4. 設定")

    def _build_files_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)

        # ---- 上半部：說明 + 檔案選擇 ----
        top = ttk.Frame(page)
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 8))
        top.columnconfigure(1, weight=1)

        intro = (
            "操作流程：\n"
            "  1. 按「選檔…」挑「地籍資料」xlsx（需含 縣市/行政區/大段/小段/地號 5 欄），\n"
            "     程式會自動下載地段代碼表並對碼，下方表格立即顯示對碼結果。\n"
            "     如果沒有檔案可挑，按下方「下載範例」可產生一份範例 xlsx。\n"
            "  2. 切到「2. 執行」分頁按「開始查詢」批次查詢，結果即時顯示。\n"
            "  3. 查完後按「2. 執行」分頁的「匯出 Excel」自行選位置存檔。"
        )
        tk.Label(
            top, text=intro, justify="left", anchor="w",
            background="#f5f8ff", relief="solid", borderwidth=1, padx=10, pady=8,
        ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        ttk.Label(top, text="地籍資料：").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(top, textvariable=self._input_path).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(top, text="選檔…", command=self._pick_input).grid(row=1, column=2, padx=4, pady=4)

        # ---- 中段：下載範例 + 重新載入 + 載入狀態 ----
        bar = ttk.Frame(page)
        bar.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
        ttk.Button(bar, text="下載範例", command=self._do_download_sample).pack(side="left")
        ttk.Button(bar, text="重新載入", command=self._do_prepare).pack(side="left", padx=(8, 0))
        self._prep_status = tk.StringVar(value="尚未載入")
        ttk.Label(bar, textvariable=self._prep_status, foreground="#1976d2").pack(side="left", padx=12)

        # ---- 下半部：預覽結果（三個子分頁）----
        preview_box = ttk.LabelFrame(page, text="對碼預覽")
        preview_box.grid(row=2, column=0, sticky="nsew", padx=4, pady=(4, 4))
        preview_box.rowconfigure(0, weight=1)
        preview_box.columnconfigure(0, weight=1)

        inner = ttk.Notebook(preview_box)
        inner.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.tree_valid = self._make_tree(inner, ("#", "縣市", "行政區", "大段", "小段", "地號", "→ city", "area", "section"))
        self.tree_nocode = self._make_tree(inner, ("#", "縣市", "行政區", "大段", "小段", "地號"))
        self.tree_bad = self._make_tree(inner, ("#", "縣市", "行政區", "大段", "小段", "地號"))
        inner.add(self.tree_valid.frame, text="可查詢")
        inner.add(self.tree_nocode.frame, text="找不到代碼")
        inner.add(self.tree_bad.frame, text="地號格式錯誤")

        # ---- 段名代碼表分頁（顯示整份地段代碼表，含連動下拉篩選）----
        sect_tab = ttk.Frame(inner)
        sect_tab.rowconfigure(1, weight=1)
        sect_tab.columnconfigure(0, weight=1)

        filter_bar = ttk.Frame(sect_tab)
        filter_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        # 三個連動下拉（可打字過濾）
        ttk.Label(filter_bar, text="縣市：").pack(side="left")
        self._sect_city = ttk.Combobox(filter_bar, width=10)
        self._sect_city.pack(side="left", padx=(0, 6))
        self._sect_city.bind("<<ComboboxSelected>>", lambda e: self._on_sect_city_change())
        self._sect_city.bind(
            "<KeyRelease>",
            lambda e: self._ac_keyrelease(self._sect_city, self._sect_city_master, e))

        ttk.Label(filter_bar, text="行政區：").pack(side="left")
        self._sect_dist = ttk.Combobox(filter_bar, width=10)
        self._sect_dist.pack(side="left", padx=(0, 6))
        self._sect_dist.bind("<<ComboboxSelected>>", lambda e: self._on_sect_dist_change())
        self._sect_dist.bind(
            "<KeyRelease>",
            lambda e: self._ac_keyrelease(self._sect_dist, self._sect_dist_master, e))

        ttk.Label(filter_bar, text="大段：").pack(side="left")
        self._sect_seg = ttk.Combobox(filter_bar, width=14)
        self._sect_seg.pack(side="left", padx=(0, 6))
        self._sect_seg.bind("<<ComboboxSelected>>", lambda e: self._refresh_section_tree())
        self._sect_seg.bind(
            "<KeyRelease>",
            lambda e: self._ac_keyrelease(self._sect_seg, self._sect_seg_master, e))

        # 關鍵字篩選框（跨所有欄位的子字串比對）
        ttk.Label(filter_bar, text="關鍵字：").pack(side="left")
        self._sect_filter = tk.StringVar()
        ent = ttk.Entry(filter_bar, textvariable=self._sect_filter, width=16)
        ent.pack(side="left", padx=(0, 6))
        ent.bind("<KeyRelease>", lambda e: self._refresh_section_tree())

        ttk.Button(filter_bar, text="清除", command=self._clear_sect_filter).pack(side="left")
        self._sect_count = tk.StringVar(value="尚未載入")
        ttk.Label(filter_bar, textvariable=self._sect_count,
                  foreground="#1976d2").pack(side="left", padx=12)

        self.tree_section = self._make_tree(
            sect_tab, ("#", "縣市", "行政區", "事務所", "大段", "小段",
                       "代碼", "所區碼", "city", "area", "office", "備註"))
        self.tree_section.frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        inner.add(sect_tab, text="段名代碼表")

        return page

    # 篩選結果最多顯示幾筆（避免 1.7 萬筆全塞進 Treeview 卡頓）
    _SECTION_TREE_LIMIT = 1000
    # 下拉選單「不篩選」的選項文字
    _SECT_ALL = "（全部）"

    # 打字時要忽略的非輸入按鍵
    _AC_SKIP_KEYS = frozenset((
        "Up", "Down", "Left", "Right", "Return", "Escape", "Tab", "Prior", "Next",
        "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
    ))

    def _ac_keyrelease(self, combo, master: list[str], event) -> None:
        """combobox 打字時即時過濾下拉清單，並重整表格。"""
        if event.keysym in self._AC_SKIP_KEYS:
            return
        typed = combo.get().strip().lower()
        if typed:
            combo["values"] = [v for v in master if typed in v.lower()] or master
        else:
            combo["values"] = master
        self._refresh_section_tree()

    def _populate_sect_filters(self) -> None:
        """段碼表載入後：縣市下拉填好、行政區/大段清空。"""
        rows = list(self._section_index.values())
        cities = sorted({r["縣市"] for r in rows if r["縣市"]})
        self._sect_city_master = [self._SECT_ALL] + cities
        self._sect_dist_master = [self._SECT_ALL]
        self._sect_seg_master = [self._SECT_ALL]
        self._sect_city["values"] = self._sect_city_master
        self._sect_city.set(self._SECT_ALL)
        self._sect_dist["values"] = self._sect_dist_master
        self._sect_dist.set(self._SECT_ALL)
        self._sect_seg["values"] = self._sect_seg_master
        self._sect_seg.set(self._SECT_ALL)

    def _on_sect_city_change(self) -> None:
        """選了縣市 → 重填行政區下拉、清空大段。"""
        city = self._sect_city.get()
        rows = list(self._section_index.values())
        if city and city != self._SECT_ALL:
            dists = sorted({r["行政區"] for r in rows if r["縣市"] == city and r["行政區"]})
        else:
            dists = []
        self._sect_dist_master = [self._SECT_ALL] + dists
        self._sect_seg_master = [self._SECT_ALL]
        self._sect_dist["values"] = self._sect_dist_master
        self._sect_dist.set(self._SECT_ALL)
        self._sect_seg["values"] = self._sect_seg_master
        self._sect_seg.set(self._SECT_ALL)
        self._refresh_section_tree()

    def _on_sect_dist_change(self) -> None:
        """選了行政區 → 重填大段下拉。"""
        city = self._sect_city.get()
        dist = self._sect_dist.get()
        rows = list(self._section_index.values())
        if (city and city != self._SECT_ALL) and (dist and dist != self._SECT_ALL):
            segs = sorted({
                r["大段"] for r in rows
                if r["縣市"] == city and r["行政區"] == dist and r["大段"]
            })
        else:
            segs = []
        self._sect_seg_master = [self._SECT_ALL] + segs
        self._sect_seg["values"] = self._sect_seg_master
        self._sect_seg.set(self._SECT_ALL)
        self._refresh_section_tree()

    def _clear_sect_filter(self) -> None:
        """清除三個下拉 + 關鍵字框，回到全部。"""
        self._sect_filter.set("")
        self._populate_sect_filters()
        self._refresh_section_tree()

    def _refresh_section_tree(self) -> None:
        """依縣市/行政區/大段下拉 + 關鍵字框重建「段名代碼表」的 Treeview。"""
        tree = self.tree_section
        tree.delete(*tree.get_children())

        rows = list(self._section_index.values())
        if not rows:
            self._sect_count.set("尚未載入（先到上方挑輸入檔）")
            return

        # 下拉用子字串比對（這樣打字打一半也能過濾）
        city = self._sect_city.get().strip()
        dist = self._sect_dist.get().strip()
        seg = self._sect_seg.get().strip()
        kw = self._sect_filter.get().strip().lower()
        if city and city != self._SECT_ALL:
            rows = [r for r in rows if city in r["縣市"]]
        if dist and dist != self._SECT_ALL:
            rows = [r for r in rows if dist in r["行政區"]]
        if seg and seg != self._SECT_ALL:
            rows = [r for r in rows if seg in r["大段"]]
        if kw:
            rows = [
                r for r in rows
                if kw in (
                    f"{r['縣市']}{r['行政區']}{r['事務所']}{r['大段']}{r['小段']}"
                    f"{r['所區碼']}{r['備註']}"
                    f"{r['city']}{r['area']}{r['office']}{r['section']}"
                ).lower()
            ]

        shown = rows[: self._SECTION_TREE_LIMIT]
        for i, r in enumerate(shown, start=1):
            tree.insert("", "end", values=(
                i, r["縣市"], r["行政區"], r["事務所"], r["大段"], r["小段"],
                r["section"], r["所區碼"], r["city"], r["area"], r["office"],
                r["備註"],
            ))

        total = len(rows)
        if total > self._SECTION_TREE_LIMIT:
            self._sect_count.set(
                f"符合 {total} 筆，顯示前 {self._SECTION_TREE_LIMIT} 筆（請縮小篩選範圍）")
        else:
            self._sect_count.set(f"符合 {total} 筆")

    @staticmethod
    def _make_tree(parent, columns):
        """建立含捲軸的 Treeview，包成 frame 回傳。"""
        frame = ttk.Frame(parent)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for c in columns:
            tree.heading(c, text=c)
            tree.column(c, width=110, anchor="w", stretch=False)
        tree.column("#", width=50, anchor="e")
        vbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hbar = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        tree.frame = frame  # 方便外面 add 進 notebook
        return tree

    # ===== API 執行分頁（唯一的執行分頁） ================================

    @staticmethod
    def _default_export_cols() -> list[dict]:
        """從 EXPORT_COLUMNS_TEMPLATE 複製出可變的設定（含 enabled=True）。"""
        return [dict(c, enabled=True) for c in EXPORT_COLUMNS_TEMPLATE]

    def _enabled_export_cols(self) -> list[dict]:
        """目前被勾選顯示的欄位（給 done tree / Excel 匯出共用）。"""
        return [c for c in self._export_cols if c.get("enabled", True)]

    # done tree 欄位 = 「#」+ 目前勾選顯示的匯出欄
    @property
    def _DONE_COLS(self) -> list[str]:
        return ["#"] + [c["name"] for c in self._enabled_export_cols()]

    # fail tree 欄位 = 「#」+「查詢狀態」+ 輸入 5 欄（簡潔，看失敗原因用）
    _FAIL_COLS = [
        "#", "查詢狀態", "輸入縣市", "輸入行政區", "輸入大段", "輸入小段", "輸入地號",
    ]
    # 各欄位寬度（沒列出的用 _DEFAULT_COL_WIDTH）
    _API_COL_WIDTHS = {
        "#": 50,
        "查詢狀態": 240,
        "輸入縣市": 80, "輸入行政區": 90, "輸入大段": 100, "輸入小段": 70, "輸入地號": 80,
        "面積": 80, "使用分區": 80, "使用地類別": 100,
        "登記日期": 150, "公告現值": 130, "公告地價": 130, "權利人類別": 100,
        "地籍連結": 320, "行政區": 200,
        "經緯度(度)": 180, "經緯度(度分秒)": 200,
        "TWD97(E)": 100, "TWD97(N)": 110, "TWD97": 160,
        "地號": 280,
        "所有權人": 130, "統一編號": 120, "所有權人類別": 100, "權利範圍類別": 100,
        "權利範圍持分_分母": 120, "權利範圍持分_分子": 120,
        "申報地價": 150, "管理者名稱": 200,
        "查詢縣市": 80, "查詢區": 70, "查詢地段": 90, "查詢地號": 90,
    }
    _DEFAULT_COL_WIDTH = 120

    def _build_api_run_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(4, weight=1)

        warn = (
            "🧪 直接打 NLSC API（不開瀏覽器）查詢。\n"
            "查詢結果分『完成』與『有問題』兩頁；「重試有問題的」會把成功的搬到『完成』。\n"
            "下方表格欄位跟匯出 Excel 完全一致；在結果列上按右鍵可看該筆的「API 呼叫明細」。\n"
            "\n"
            "☐ 歷年國土利用：勾選才會去抓「國土利用_年月」/「國土利用_現況」兩個欄位。\n"
            "　　不需要就不需勾選。（會另外抓取歷年國土利用的 API）"
        )
        tk.Label(
            page, text=warn, justify="left", anchor="w",
            background="#e3f2fd", relief="solid", borderwidth=1, padx=10, pady=8,
        ).grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 6))

        bar = ttk.Frame(page)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.btn_start_api = ttk.Button(bar, text="開始查詢", command=self._do_run_api)
        self.btn_start_api.pack(side="left")
        self.btn_stop_api = ttk.Button(bar, text="停止", command=self._do_stop_api, state="disabled")
        self.btn_stop_api.pack(side="left", padx=8)
        self.btn_retry_api = ttk.Button(bar, text="重試有問題的", command=self._do_retry_failed_api, state="disabled")
        self.btn_retry_api.pack(side="left", padx=(0, 12))
        self.btn_clear_api = ttk.Button(bar, text="清空結果", command=self._do_clear_results_api, state="disabled")
        self.btn_clear_api.pack(side="left")
        # 勾了才打第 4 顆 API（LandUsePointYears）。預設不勾 —— 多數查詢用不到這兩欄，
        # 每筆可省約 0.3 秒（整批約快一倍）；沒勾時「國土利用_年月 / _現況」兩欄留白。
        self._land_use_var = tk.BooleanVar(value=False)
        self.chk_land_use_api = ttk.Checkbutton(
            bar, text="歷年國土利用", variable=self._land_use_var)
        self.chk_land_use_api.pack(side="left", padx=(12, 0))
        self._run_status_api = tk.StringVar(value="待命")
        ttk.Label(bar, textvariable=self._run_status_api).pack(side="left", padx=12)
        # 匯出按鈕 + 欄位設定推到最右
        self.btn_export_api = ttk.Button(bar, text="匯出 Excel", command=self._do_export_api, state="disabled")
        self.btn_export_api.pack(side="right")
        ttk.Button(bar, text="欄位設定…",
                   command=self._open_export_cols_dialog).pack(side="right", padx=(0, 8))

        ttk.Label(page, text="進度：").grid(row=2, column=0, sticky="w", padx=4)
        self.progress_api = ttk.Progressbar(page, mode="determinate", maximum=100)
        self.progress_api.grid(row=3, column=0, sticky="ew", padx=4, pady=(0, 8))

        # 兩個分頁：完成 / 有問題
        self.api_result_nb = ttk.Notebook(page)
        self.api_result_nb.grid(row=4, column=0, sticky="nsew", padx=4, pady=(4, 4))

        done_frame, self.result_tree_api_done = self._build_result_tree(self.api_result_nb, "done")
        fail_frame, self.result_tree_api_fail = self._build_result_tree(self.api_result_nb, "fail")
        self.api_result_nb.add(done_frame, text="✓ 完成 (0)")
        self.api_result_nb.add(fail_frame, text="✗ 有問題 (0)")

        return page

    def _build_result_tree(self, parent, kind: str) -> tuple[ttk.Frame, ttk.Treeview]:
        """建一個含捲軸的 Treeview。kind='done' 用 _DONE_COLS，'fail' 用 _FAIL_COLS。"""
        frame = ttk.Frame(parent)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        cols = self._DONE_COLS if kind == "done" else self._FAIL_COLS
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=14)
        self._apply_api_column_widths(tree)
        vbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hbar = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        if kind == "fail":
            tree.tag_configure("row", background="#fdecea")
        # 右鍵 → 彈視窗顯示這筆的 4 顆 API 原始回應
        tree.bind("<Button-3>", lambda e, k=kind: self._on_result_right_click(e, k))
        return frame, tree

    def _on_result_right_click(self, event, kind: str) -> None:
        """右鍵按到某列：彈視窗顯示該筆的原始 API 回應。"""
        tree = event.widget
        row_iid = tree.identify_row(event.y)
        if not row_iid:
            return
        try:
            idx = int(row_iid)
        except ValueError:
            return
        results = (self._results_api_done if kind == "done"
                   else self._results_api_fail)
        if 0 < idx <= len(results):
            self._show_raw_api_response(results[idx - 1])

    # 每顆 API 的說明，給彈窗標題用
    _API_NOTES = {
        "getLandInfoSect":   "土地基本資訊 + 所有人 + 公有土地",
        "qryTileMapIndex":   "地塊中心經緯度 + 地段中文名",
        "LocationQuery":     "行政區 + 經緯度(度/度分秒) + 國土利用",
        "LandUsePointYears": "歷年國土利用調查（匯出的是最新一期）",
    }

    @staticmethod
    def _format_trace_body(data: dict, entry: dict) -> str:
        """把一筆 trace 的回應內容取出來（從 data 的 _raw_* 欄位）。"""
        raw = data.get(entry.get("raw_key") or "")
        if raw is None or raw == "":
            return ""
        if isinstance(raw, (dict, list)):
            return json.dumps(raw, ensure_ascii=False, indent=2)
        return str(raw)

    def _show_raw_api_response(self, data: dict) -> None:
        """彈出視窗顯示這筆資料每顆 API 的請求參數 + 原始回應。"""
        top = tk.Toplevel(self.root)
        top.title("API 呼叫明細")
        top.geometry("960x760")

        head = (
            f"{data.get('輸入縣市','')} {data.get('輸入行政區','')} "
            f"{data.get('輸入大段','')}{data.get('輸入小段','')} "
            f"{data.get('輸入地號','')}    [{data.get('查詢狀態','')}]"
        )
        tk.Label(top, text=head, anchor="w", justify="left",
                 background="#e3f2fd", padx=10, pady=6,
                 ).pack(fill="x", padx=8, pady=(8, 4))

        txt = ScrolledText(top, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=8, pady=4)

        trace = data.get("_api_trace")
        if not trace:
            # 舊結果（這版之前查的）沒有 trace，退回只顯示回應
            txt.insert("end", "（這筆沒有呼叫紀錄，可能是舊版查詢的結果；只顯示回應）\n\n")
            for title, key in (
                ("1. getLandInfoSect", "_raw_payload"),
                ("2. qryTileMapIndex", "_raw_tile"),
                ("3. LocationQuery", "_raw_location"),
                ("4. LandUsePointYears", "_raw_land_use"),
            ):
                txt.insert("end", "=" * 74 + "\n" + title + "\n" + "=" * 74 + "\n")
                body = self._format_trace_body(data, {"raw_key": key})
                txt.insert("end", (body or "(無資料)") + "\n\n")
            txt.configure(state="disabled")
        else:
            n_called = sum(1 for e in trace if not e.get("skipped"))
            total_ms = sum(e.get("elapsed") or 0 for e in trace)
            txt.insert("end", f"這筆共呼叫 {n_called} 顆 API，合計 {total_ms:.2f} 秒\n\n")

            for e in trace:
                seq = e.get("seq", "?")
                name = e.get("name", "?")
                note_of = self._API_NOTES.get(name, "")

                # 標題列：狀態圖示 + HTTP 狀態 + 耗時 + 大小
                if e.get("skipped"):
                    mark = "⊘ 跳過"
                elif e.get("error"):
                    mark = "✗ 失敗"
                else:
                    mark = "✓ 成功"
                bits = []
                if e.get("status") is not None:
                    bits.append(f"HTTP {e['status']}")
                if e.get("elapsed") is not None:
                    bits.append(f"{e['elapsed']:.2f} 秒")
                if e.get("bytes") is not None:
                    bits.append(f"{e['bytes']} bytes")
                if e.get("attempts", 1) > 1:
                    bits.append(f"打了 {e['attempts']} 次")
                tail = ("  (" + ", ".join(bits) + ")") if bits else ""

                txt.insert("end", "=" * 74 + "\n")
                txt.insert("end", f"{seq}. {name}  {mark}{tail}\n")
                if note_of:
                    txt.insert("end", f"   {note_of}\n")
                txt.insert("end", "=" * 74 + "\n")

                if e.get("skipped"):
                    txt.insert("end", (e.get("note") or "(跳過)") + "\n\n")
                    continue

                txt.insert("end", f"[請求] {e.get('method','')} {e.get('url','')}\n")
                params = e.get("params")
                if params:
                    pad = " " * 7
                    lines = [f"{k} = {v}" for k, v in params.items()]
                    txt.insert("end", "[參數] " + f"\n{pad}".join(lines) + "\n")
                if e.get("note"):
                    txt.insert("end", f"[備註] {e['note']}\n")
                if e.get("error"):
                    txt.insert("end", f"[錯誤] {e['error']}\n")

                body = self._format_trace_body(data, e)
                txt.insert("end", "[回應]\n")
                txt.insert("end", (body or "(空)") + "\n\n")

            txt.configure(state="disabled")

        bottom = ttk.Frame(top)
        bottom.pack(fill="x", padx=8, pady=(0, 8))

        def copy_all():
            top.clipboard_clear()
            top.clipboard_append(txt.get("1.0", "end-1c"))
            messagebox.showinfo("已複製", "整份明細已複製到剪貼簿。", parent=top)

        ttk.Button(bottom, text="複製全部", command=copy_all).pack(side="left")
        ttk.Button(bottom, text="關閉", command=top.destroy).pack(side="right")
        ttk.Button(
            bottom, text="複製全部",
            command=lambda: (top.clipboard_clear(), top.clipboard_append(txt.get("1.0", "end-1c"))),
        ).pack(side="right", padx=(0, 8))

    def _apply_api_column_widths(self, tree: ttk.Treeview) -> None:
        """套用欄寬。"""
        for c in list(tree["columns"]):
            tree.heading(c, text=c)
            w = self._API_COL_WIDTHS.get(c, self._DEFAULT_COL_WIDTH)
            anchor = "e" if c == "#" else "w"
            tree.column(c, width=w, anchor=anchor, stretch=False)

    # ===== 匯出欄位設定 =====================================================

    @staticmethod
    def _excel_col_letter(n: int) -> str:
        """1→A, 26→Z, 27→AA, 30→AD"""
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def _apply_export_cols_change(self) -> None:
        """欄位設定變更後：done tree 重建欄位 + 重新填值。"""
        cols = self._DONE_COLS
        tree = self.result_tree_api_done
        tree["columns"] = cols
        # heading + 寬度重套
        self._apply_api_column_widths(tree)
        # 重新塞所有目前完成的結果
        tree.delete(*tree.get_children())
        for i, r in enumerate(self._results_api_done, start=1):
            self._insert_done_row(tree, r, i)

    def _open_export_cols_dialog(self) -> None:
        """彈出「匯出欄位設定」對話框。"""
        top = tk.Toplevel(self.root)
        top.title("匯出欄位設定")
        top.geometry("1020x720")
        top.transient(self.root)

        def _sort_unchecked_to_bottom(cols: list[dict]) -> list[dict]:
            """穩定排序：未勾選的搬到最下方，同組內保留原順序。"""
            return sorted(cols, key=lambda c: 0 if c.get("enabled", True) else 1)

        # 暫存工作副本，按「套用」才寫回 self._export_cols
        # 開啟時穩定排序：未勾選的移到最下方（保留同組內原本順序）
        working: list[dict] = _sort_unchecked_to_bottom([dict(c) for c in self._export_cols])

        # ====== 設定檔列：下拉 / 另存 / 刪除 ===================================
        preset_bar = ttk.Frame(top)
        preset_bar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(preset_bar, text="設定檔：").pack(side="left")
        preset_var = tk.StringVar(value="")
        preset_cb = ttk.Combobox(
            preset_bar, textvariable=preset_var, width=22, state="readonly")
        preset_cb.pack(side="left", padx=(0, 6))

        DEFAULT_LABEL = "（預設）"

        def reload_preset_list(select: str | None = None) -> None:
            names = list_export_presets()
            preset_cb["values"] = [DEFAULT_LABEL] + names
            if select and select in names:
                preset_var.set(select)
            elif preset_var.get() not in preset_cb["values"]:
                preset_var.set(DEFAULT_LABEL)

        # 啟動時：把下拉值設成上次套用的 preset 名（若有）
        initial_name = load_active_preset_name()
        reload_preset_list(select=initial_name if initial_name else None)
        if not initial_name or initial_name not in list_export_presets():
            preset_var.set(DEFAULT_LABEL)

        def on_preset_pick(_e=None):
            name = preset_var.get()
            if name == DEFAULT_LABEL:
                # 選「預設」= 載入原始預設值
                working.clear()
                working.extend(_sort_unchecked_to_bottom(self._default_export_cols()))
                refresh()
                return
            if not name:
                return
            cols = load_export_preset(name)
            if cols is None:
                messagebox.showwarning("讀取失敗", f"設定檔「{name}」讀不到。")
                return
            working.clear()
            working.extend(_sort_unchecked_to_bottom(cols))
            refresh()

        preset_cb.bind("<<ComboboxSelected>>", on_preset_pick)

        def do_save_as():
            name = simpledialog.askstring(
                "另存設定檔", "請輸入設定檔名稱：", parent=top)
            if name is None:
                return
            name = name.strip()
            if not name:
                messagebox.showwarning("名稱不可空白", "請輸入名稱。", parent=top)
                return
            if name in list_export_presets():
                if not messagebox.askyesno(
                        "確認覆蓋", f"設定檔「{name}」已存在，要覆蓋嗎？", parent=top):
                    return
            if save_export_preset(name, working):
                reload_preset_list(select=name)
                messagebox.showinfo("已存", f"設定檔「{name}」已儲存。", parent=top)
            else:
                messagebox.showerror("儲存失敗", "寫入 export_cols.json 失敗。", parent=top)

        def do_delete_preset():
            name = preset_var.get()
            if not name or name == DEFAULT_LABEL:
                messagebox.showinfo("沒選", "請先在下拉選一個要刪除的設定檔。", parent=top)
                return
            if not messagebox.askyesno("確認刪除", f"刪除設定檔「{name}」？", parent=top):
                return
            if delete_export_preset(name):
                reload_preset_list()
            else:
                messagebox.showerror("刪除失敗", "寫入 export_cols.json 失敗。", parent=top)

        ttk.Button(preset_bar, text="另存…", command=do_save_as).pack(side="left", padx=(0, 4))
        ttk.Button(preset_bar, text="刪除", command=do_delete_preset).pack(side="left")

        # ====== 操作說明 ======================================================
        tk.Label(
            top, anchor="w", justify="left", padx=10, pady=8,
            background="#f5f8ff", relief="solid", borderwidth=1,
            text=("• 從上方「設定檔」下拉可載入已存好的組合（套用前在此預覽）\n"
                  "• 點「顯示」欄的 ☑/☐ 切換是否匯出；雙擊「欄位名」可改名\n"
                  "• 拖曳列可調整順序;「Excel 欄」會即時更新"),
        ).pack(fill="x", padx=8, pady=(0, 4))

        # ====== Treeview ======================================================
        mid = ttk.Frame(top)
        mid.pack(fill="both", expand=True, padx=8, pady=4)
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        tree = ttk.Treeview(
            mid, columns=("顯示", "Excel欄", "欄位名", "資料來源", "處理"),
            show="headings", selectmode="browse",
        )
        widths = {"顯示": 50, "Excel欄": 60, "欄位名": 170, "資料來源": 230, "處理": 380}
        for c in tree["columns"]:
            tree.heading(c, text=c)
            tree.column(c, width=widths[c], anchor="w", stretch=(c == "欄位名"))
        tree.column("顯示", anchor="center")
        tree.column("Excel欄", anchor="center")
        tree.tag_configure("drag_hover", background="#ffe0b2")  # 拖曳目標高亮

        vbar = ttk.Scrollbar(mid, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")

        def refresh(select_iid: str | None = None) -> None:
            tree.delete(*tree.get_children())
            shown_pos = 0
            for i, c in enumerate(working):
                enabled = c.get("enabled", True)
                if enabled:
                    shown_pos += 1
                    letter = self._excel_col_letter(shown_pos)
                else:
                    letter = ""  # 未勾選 → 空白（不顯示「—」）
                desc_src, desc_tx = COLUMN_DESCRIPTIONS.get(
                    (c.get("source"), c.get("transform")),
                    (c.get("source") or "", c.get("transform") or ""),  # 沒對應就退回原始值
                )
                tree.insert(
                    "", "end", iid=str(i),
                    values=(
                        "☑" if enabled else "☐",
                        letter,
                        c["name"],
                        desc_src,
                        desc_tx,
                    ),
                )
            if select_iid is not None and select_iid in tree.get_children():
                tree.selection_set(select_iid)
                tree.see(select_iid)

        refresh()

        # ---- 拖曳排序 + 點 ☑/☐ 切換 ----
        drag = {"src": None}  # type: dict[str, str | None]

        def _clear_hover():
            for iid in tree.get_children():
                tags = list(tree.item(iid, "tags"))
                if "drag_hover" in tags:
                    tags.remove("drag_hover")
                    tree.item(iid, tags=tags)

        def on_press(event):
            if tree.identify_region(event.x, event.y) != "cell":
                return
            col = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
            if not row:
                return
            idx = int(row)
            if col == "#1":  # 顯示欄 → 切換 enabled，不啟動拖曳
                working[idx]["enabled"] = not working[idx].get("enabled", True)
                refresh(select_iid=row)
                return
            # 其他欄：標記拖曳來源
            drag["src"] = row
            tree.selection_set(row)

        def on_motion(event):
            if drag["src"] is None:
                return
            target = tree.identify_row(event.y)
            _clear_hover()
            if target and target != drag["src"]:
                tags = list(tree.item(target, "tags"))
                tags.append("drag_hover")
                tree.item(target, tags=tags)

        def on_release(event):
            src_iid = drag["src"]
            drag["src"] = None
            _clear_hover()
            if src_iid is None:
                return
            target_iid = tree.identify_row(event.y)
            if not target_iid or target_iid == src_iid:
                return
            src_idx = int(src_iid)
            target_idx = int(target_iid)
            item = working.pop(src_idx)
            # pop 後，若 src < target 則 target 在 list 中往前移 1
            if src_idx < target_idx:
                target_idx -= 1
            working.insert(target_idx, item)
            refresh(select_iid=str(target_idx))

        tree.bind("<ButtonPress-1>", on_press)
        tree.bind("<B1-Motion>", on_motion)
        tree.bind("<ButtonRelease-1>", on_release)

        # ---- 雙擊「欄位名」改名（Entry overlay）----
        def on_double_click(event):
            if tree.identify_region(event.x, event.y) != "cell":
                return
            if tree.identify_column(event.x) != "#3":
                return
            row = tree.identify_row(event.y)
            if not row:
                return
            idx = int(row)
            bbox = tree.bbox(row, "#3")
            if not bbox:
                return
            x, y, w, h = bbox
            entry = ttk.Entry(tree)
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, working[idx]["name"])
            entry.select_range(0, "end")
            entry.focus_set()

            def commit(_e=None):
                new_name = entry.get().strip()
                if new_name:
                    working[idx]["name"] = new_name
                entry.destroy()
                refresh(select_iid=row)

            entry.bind("<Return>", commit)
            entry.bind("<FocusOut>", commit)
            entry.bind("<Escape>", lambda e: entry.destroy())

        tree.bind("<Double-Button-1>", on_double_click)

        # ====== 控制列：全選 / 全不選 / 恢復預設 ==============================
        ctl = ttk.Frame(top)
        ctl.pack(fill="x", padx=8, pady=4)

        def do_check_all():
            for c in working:
                c["enabled"] = True
            refresh()

        def do_uncheck_all():
            for c in working:
                c["enabled"] = False
            refresh()

        def do_reset():
            working.clear()
            working.extend(self._default_export_cols())
            refresh()

        ttk.Button(ctl, text="全選 ☑", command=do_check_all).pack(side="left")
        ttk.Button(ctl, text="全不選 ☐", command=do_uncheck_all).pack(side="left", padx=(4, 0))
        ttk.Button(ctl, text="恢復預設", command=do_reset).pack(side="left", padx=(16, 0))

        # ====== 套用 / 取消 ===================================================
        bottom = ttk.Frame(top)
        bottom.pack(fill="x", padx=8, pady=(0, 8))

        def apply_and_close():
            self._export_cols = [dict(c) for c in working]
            self._apply_export_cols_change()
            # 紀錄目前下拉選的 preset 名（讓下次開啟還記得）
            current_pick = preset_var.get()
            preset_to_save = "" if current_pick == DEFAULT_LABEL else current_pick
            if not save_export_cols(self._export_cols, preset_name=preset_to_save):
                messagebox.showwarning(
                    "設定無法儲存",
                    "欄位設定已套用，但寫入 export_cols.json 失敗，下次開啟會回到預設。",
                )
            top.destroy()

        ttk.Button(bottom, text="取消", command=top.destroy).pack(side="right")
        ttk.Button(bottom, text="套用", command=apply_and_close).pack(side="right", padx=(0, 8))

    @staticmethod
    def _is_success(data: dict) -> bool:
        status = str(data.get("查詢狀態", "")).strip()
        return status == "" or status == "成功"

    def _transform_for_template(self, data: dict) -> dict:
        """套用目前勾選顯示的欄位設定 + transform，得到顯示/匯出用的 dict。"""
        out = {}
        for spec in self._enabled_export_cols():
            name = spec["name"]
            src = spec.get("source")
            tx = spec.get("transform")
            if src is None:
                out[name] = ""
                continue
            val = data.get(src, "")
            if val is None:
                val = ""
            if tx and val != "":
                fn = _TRANSFORMS.get(tx)
                if fn is not None:
                    try:
                        val = fn(val)
                    except Exception:
                        pass
            out[name] = val
        return out

    def _refresh_api_tab_counts(self) -> None:
        self.api_result_nb.tab(0, text=f"✓ 完成 ({len(self._results_api_done)})")
        self.api_result_nb.tab(1, text=f"✗ 有問題 ({len(self._results_api_fail)})")

    def _refresh_api_buttons(self) -> None:
        has_done = bool(self._results_api_done)
        has_fail = bool(self._results_api_fail)
        has_any = has_done or has_fail
        self.btn_clear_api.configure(state="normal" if has_any else "disabled")
        self.btn_export_api.configure(state="normal" if has_done else "disabled")
        self.btn_retry_api.configure(state="normal" if (has_fail and not self._running_api) else "disabled")

    def _insert_done_row(self, tree: ttk.Treeview, data: dict, idx: int) -> None:
        """done tree 插入：對 data 做 template transform 後再填值。"""
        view = self._transform_for_template(data)
        cols = list(tree["columns"])
        values = []
        for c in cols:
            if c == "#":
                values.append(idx)
            else:
                v = view.get(c, "")
                values.append("" if v is None else str(v))
        # iid = 字串化的 idx，給右鍵選單反查 _results_api_done 用
        tree.insert("", "end", iid=str(idx), values=values)

    def _insert_fail_row(self, tree: ttk.Treeview, data: dict, idx: int) -> None:
        cols = list(tree["columns"])
        values = []
        for c in cols:
            if c == "#":
                values.append(idx)
            else:
                v = data.get(c, "")
                values.append("" if v is None else str(v))
        tree.insert("", "end", iid=str(idx), values=values, tags=("row",))

    def _refill_done_tree(self) -> None:
        tree = self.result_tree_api_done
        tree.delete(*tree.get_children())
        for i, r in enumerate(self._results_api_done, start=1):
            self._insert_done_row(tree, r, i)

    def _refill_fail_tree(self) -> None:
        tree = self.result_tree_api_fail
        tree.delete(*tree.get_children())
        for i, r in enumerate(self._results_api_fail, start=1):
            self._insert_fail_row(tree, r, i)

    def _append_result_row_api(self, data: dict) -> None:
        """收到一筆查詢結果：依狀態插到對應的 tree。"""
        if self._is_success(data):
            self._results_api_done.append(data)
            self._insert_done_row(self.result_tree_api_done, data, len(self._results_api_done))
        else:
            self._results_api_fail.append(data)
            self._insert_fail_row(self.result_tree_api_fail, data, len(self._results_api_fail))
        self._refresh_api_tab_counts()
        self._refresh_api_buttons()

    def _do_run_api(self, rows: list | None = None, is_retry: bool = False) -> None:
        """rows=None 表示跑 prepared.valid 全部；否則跑指定那些（給 retry 用）。"""
        if self._running_api:
            return
        if rows is None:
            if not self._prepared or not self._prepared.valid:
                messagebox.showwarning("沒有資料", "請先到『檔案與預覽』分頁挑輸入檔（會自動載入）")
                self.notebook.select(0)
                return
            if self._results_api_done or self._results_api_fail:
                ans = messagebox.askyesnocancel(
                    "已有查詢結果",
                    f"目前有 {len(self._results_api_done)} 筆成功、{len(self._results_api_fail)} 筆問題。\n"
                    "是 = 清空後重新查詢\n"
                    "否 = 保留並把新結果附加在後\n"
                    "取消 = 不執行")
                if ans is None:
                    return
                if ans:
                    self._do_clear_results_api()
            rows = list(self._prepared.valid)

        self._running_api = True
        self._stop_flag_api = False
        # worker 是另一條 thread，不能直接讀 tk 變數 → 先在主 thread 取值
        enable_land_use = bool(self._land_use_var.get())
        self._log(f"[API] 歷年國土利用：{'要查' if enable_land_use else '不查（跳過第 4 顆 API）'}")
        self.btn_start_api.configure(state="disabled")
        self.chk_land_use_api.configure(state="disabled")
        self.btn_stop_api.configure(state="normal")
        self.btn_retry_api.configure(state="disabled")
        self.btn_export_api.configure(state="disabled")
        self.progress_api.configure(value=0, maximum=max(1, len(rows)))
        self._run_status_api.set(f"{'重試' if is_retry else '執行'}中… 0/{len(rows)}")

        def worker():
            try:
                run_api_query(
                    rows, self.cfg,
                    log=lambda m: self.root.after(0, lambda m=m: self._log(m)),
                    progress=lambda i, n: self.root.after(0, lambda i=i, n=n: self._on_progress_api(i, n)),
                    should_stop=lambda: self._stop_flag_api,
                    on_row=lambda d: self.root.after(0, lambda d=d: self._append_result_row_api(d)),
                    enable_land_use=enable_land_use,
                )
                self.root.after(0, self._on_run_done_api)
            except StopRequested:
                self.root.after(0, self._on_run_stopped_api)
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._on_run_failed_api(msg))

        self._worker_api = threading.Thread(target=worker, daemon=True)
        self._worker_api.start()

    def _do_stop_api(self) -> None:
        if not self._running_api:
            return
        self._stop_flag_api = True
        self._run_status_api.set("等待本筆結束…")
        self._log("使用者按下停止；本筆結束後中止")

    def _do_retry_failed_api(self) -> None:
        """重試『有問題』那邊的查詢；成功的搬到『完成』，失敗的留在『有問題』。"""
        if self._running_api:
            return
        if not self._results_api_fail:
            messagebox.showinfo("沒有資料", "『有問題』那邊沒有資料可以重試")
            return
        if not self._prepared:
            messagebox.showwarning("缺少資料", "原始查詢條件已遺失，請先到『預覽』重新載入")
            return
        fail_keys = set()
        for r in self._results_api_fail:
            key = (r.get("輸入縣市", ""), r.get("輸入行政區", ""),
                   r.get("輸入大段", ""), r.get("輸入小段", ""), r.get("輸入地號", ""))
            fail_keys.add(key)
        retry_rows = [
            p for p in self._prepared.valid
            if (p.輸入縣市, p.輸入行政區, p.輸入大段, p.輸入小段, p.輸入地號) in fail_keys
        ]
        if not retry_rows:
            messagebox.showwarning("找不到原始條件",
                "對應的原始 PreparedRow 找不到，可能是匯入後重新對碼造成。"
                "請按「清空結果」後重新「開始查詢」。")
            return
        self._results_api_fail.clear()
        self._refill_fail_tree()
        self._refresh_api_tab_counts()
        self._refresh_api_buttons()
        self._log(f"重試『有問題』{len(retry_rows)} 筆…")
        self._do_run_api(rows=retry_rows, is_retry=True)

    def _do_clear_results_api(self) -> None:
        self._results_api_done.clear()
        self._results_api_fail.clear()
        self._refill_done_tree()
        self._refill_fail_tree()
        self._refresh_api_tab_counts()
        self._refresh_api_buttons()

    def _do_export_api(self) -> None:
        """匯出『完成』那邊資料（套 EXPORT_COLUMNS_TEMPLATE 處理）。"""
        if not self._results_api_done:
            messagebox.showinfo("沒有結果", "『完成』沒有可匯出的結果")
            return
        start = str(Path(self._output_path.get()).parent) if self._output_path.get() else ""
        name = Path(self._output_path.get()).stem + ".xlsx"
        path = filedialog.asksaveasfilename(
            title="匯出 Excel 至", initialdir=start, defaultextension=".xlsx",
            initialfile=name, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            export_results_template(self._results_api_done, path,
                                     columns=self._enabled_export_cols())
            n = len(self._results_api_done)
            self._log(f"已匯出 {n} 筆至 {path}")
            messagebox.showinfo("匯出完成", f"已匯出 {n} 筆至：\n{path}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

    def _on_progress_api(self, i: int, n: int) -> None:
        self.progress_api.configure(value=i, maximum=max(1, n))
        self._run_status_api.set(f"執行中 {i}/{n}")

    def _on_run_done_api(self) -> None:
        self._running_api = False
        self.btn_start_api.configure(state="normal")
        self.chk_land_use_api.configure(state="normal")
        self.btn_stop_api.configure(state="disabled")
        n_done = len(self._results_api_done)
        n_fail = len(self._results_api_fail)
        self._run_status_api.set(f"完成：成功 {n_done} 筆 / 問題 {n_fail} 筆")
        self._log(f"[API] 完成：成功 {n_done} 筆、有問題 {n_fail} 筆")
        self._refresh_api_buttons()

    def _on_run_stopped_api(self) -> None:
        self._running_api = False
        self.btn_start_api.configure(state="normal")
        self.chk_land_use_api.configure(state="normal")
        self.btn_stop_api.configure(state="disabled")
        n_done = len(self._results_api_done)
        n_fail = len(self._results_api_fail)
        self._run_status_api.set(f"已停止（成功 {n_done} 筆 / 問題 {n_fail} 筆）")
        self._log("[API] 查詢已停止")
        self._refresh_api_buttons()

    def _on_run_failed_api(self, msg: str) -> None:
        self._running_api = False
        self.btn_start_api.configure(state="normal")
        self.chk_land_use_api.configure(state="normal")
        self.btn_stop_api.configure(state="disabled")
        self._run_status_api.set("執行失敗")
        self._log(f"[API] 執行失敗：{msg.splitlines()[0]}")
        self._refresh_api_buttons()

    def _build_log_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent)
        page.rowconfigure(0, weight=1)
        page.columnconfigure(0, weight=1)
        self.log_text = ScrolledText(page, wrap="word", state="disabled", height=20)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        btn_row = ttk.Frame(page)
        btn_row.grid(row=1, column=0, sticky="e", padx=4, pady=(0, 4))
        ttk.Button(btn_row, text="清除日誌", command=self._clear_log).pack(side="right")
        return page

    def _build_settings_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent)
        page.columnconfigure(1, weight=1)

        ttk.Label(page, text=f"設定檔位置：{config_path()}", foreground="#666").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=4, pady=(4, 12))

        self._cfg_vars = {}
        defs = [
            ("section_url", "地段代碼下載 URL", "str"),
            ("landno_pattern", "地號格式 (regex)", "str"),
            ("api_land_info_url", "API: 土地資訊 URL", "str"),
            ("api_tile_index_url", "API: 圖磚索引 URL", "str"),
            ("api_location_query_url", "API: 行政區查詢 URL", "str"),
            ("api_land_use_url", "API: 歷年國土利用 URL", "str"),
            ("api_request_timeout", "API 等待秒數", "int"),
            ("api_request_delay", "每筆之間延遲秒數", "float"),
        ]
        for i, (key, label, _kind) in enumerate(defs, start=1):
            ttk.Label(page, text=label).grid(row=i, column=0, sticky="w", padx=4, pady=3)
            var = tk.StringVar(value=str(self.cfg.get(key, "")))
            self._cfg_vars[key] = var
            ttk.Entry(page, textvariable=var, width=80).grid(
                row=i, column=1, columnspan=2, sticky="ew", padx=4, pady=3)

        i = len(defs) + 1
        btn = ttk.Frame(page)
        btn.grid(row=i, column=0, columnspan=3, sticky="e", padx=4, pady=12)
        ttk.Button(btn, text="還原預設", command=self._reset_settings).pack(side="right", padx=(8, 0))
        ttk.Button(btn, text="儲存", command=self._save_settings).pack(side="right")

        return page

    # ---- 檔案分頁 ----

    def _do_download_sample(self) -> None:
        """產生範例 input.xlsx（資料寫在程式碼裡的 SAMPLE_INPUT_ROWS）。"""
        path = filedialog.asksaveasfilename(
            title="儲存範例 input.xlsx", defaultextension=".xlsx",
            initialfile="input_sample.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(INPUT_COLUMNS)
            for r in SAMPLE_INPUT_ROWS:
                ws.append(list(r))
            wb.save(path)
            wb.close()
            self._log(f"已產生範例 input.xlsx：{path}")
            ans = messagebox.askyesno(
                "範例已產生",
                f"已存至：\n{path}\n\n要直接載入這個範例檔嗎？")
            if ans:
                self._input_path.set(path)
                self._do_prepare()
        except Exception as e:
            messagebox.showerror("產生失敗", str(e))

    def _pick_input(self) -> None:
        start = str(Path(self._input_path.get()).parent) if self._input_path.get() else ""
        p = filedialog.askopenfilename(
            title="選擇 input.xlsx", initialdir=start,
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有檔案", "*.*")])
        if p:
            self._input_path.set(p)
            # 自動下載地段代碼表並對碼
            self._do_prepare()

    # ---- 載入並對碼 ----

    def _do_prepare(self) -> None:
        if self._running_api:
            messagebox.showinfo("處理中", "查詢進行中，請先停止")
            return
        path = self._input_path.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("找不到檔案", f"輸入檔不存在：\n{path}")
            return

        self._prep_status.set("下載地段代碼表中…")
        self._log("開始下載地段代碼表…")
        self.root.update_idletasks()

        def worker():
            try:
                section_index = download_section_table(self.cfg["section_url"])
                _, input_rows = _read_xlsx_rows(path)
                prepared = prepare_input(
                    input_rows, section_index, self.cfg.get("landno_pattern", DEFAULT_CONFIG["landno_pattern"]))
                self.root.after(0, lambda: self._on_prepared(prepared, section_index))
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._on_prepare_failed(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_prepared(self, prepared: PreparedInput, section_index: dict[str, dict]) -> None:
        self._prepared = prepared
        self._section_index = section_index
        self._populate_sect_filters()
        self._refresh_section_tree()
        for t in (self.tree_valid, self.tree_nocode, self.tree_bad):
            t.delete(*t.get_children())
        for i, r in enumerate(prepared.valid, start=1):
            self.tree_valid.insert("", "end", values=(
                i, r.輸入縣市, r.輸入行政區, r.輸入大段, r.輸入小段, r.landno,
                r.city, r.area, r.section,
            ))
        for i, r in enumerate(prepared.no_code, start=1):
            self.tree_nocode.insert("", "end", values=(
                i, r["輸入縣市"], r["輸入行政區"], r["輸入大段"], r["輸入小段"], r["landno"],
            ))
        for i, r in enumerate(prepared.bad_landno, start=1):
            self.tree_bad.insert("", "end", values=(
                i, r["輸入縣市"], r["輸入行政區"], r["輸入大段"], r["輸入小段"], r["landno"],
            ))
        self._prep_status.set(
            f"可查詢 {len(prepared.valid)} 筆、找不到代碼 {len(prepared.no_code)} 筆、"
            f"地號格式錯誤 {len(prepared.bad_landno)} 筆")
        self._log(self._prep_status.get())

    def _on_prepare_failed(self, msg: str) -> None:
        self._prep_status.set("載入失敗")
        self._log(f"載入失敗：{msg.splitlines()[0]}")
        messagebox.showerror("載入失敗", msg)

    # ---- 日誌 ----

    def _log(self, msg: str) -> None:
        t = datetime.now().strftime("%H:%M:%S")
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{t}] {msg}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except tk.TclError:
            pass

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ---- 設定 ----

    def _save_settings(self) -> None:
        new_cfg = dict(self.cfg)
        for key, var in self._cfg_vars.items():
            raw = var.get().strip()
            if key == "api_request_timeout":
                try:
                    new_cfg[key] = int(raw)
                except ValueError:
                    messagebox.showwarning("數值錯誤", f"{key} 必須是整數")
                    return
            elif key == "api_request_delay":
                try:
                    new_cfg[key] = float(raw)
                except ValueError:
                    messagebox.showwarning("數值錯誤", f"{key} 必須是數字")
                    return
            else:
                new_cfg[key] = raw
        try:
            save_config(new_cfg)
            self.cfg = new_cfg
            self._log(f"設定已儲存：{config_path()}")
            messagebox.showinfo("已儲存", f"設定已寫入：\n{config_path()}")
        except Exception as e:
            messagebox.showerror("存檔失敗", str(e))

    def _reset_settings(self) -> None:
        if not messagebox.askyesno("確認", "還原所有設定為預設值？"):
            return
        for key, var in self._cfg_vars.items():
            var.set(str(DEFAULT_CONFIG.get(key, "")))
        messagebox.showerror("載入失敗", msg)

    # ---- 執行分頁 ----

# ===========================================================================
# 獨立執行入口
# ===========================================================================

def main() -> int:
    root = tk.Tk()
    root.title(APP_TITLE)
    _configure_global_fonts()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    # 先設「正常尺寸」並放在主螢幕中央 — 之後從最大化還原時會回到這個尺寸，
    # 避免直接 state("zoomed") 啟動後還原成橫跨多螢幕的長條視窗。
    NORMAL_W, NORMAL_H = 1400, 850
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = max(0, (sw - NORMAL_W) // 2)
    y = max(0, (sh - NORMAL_H) // 2 - 30)
    root.geometry(f"{NORMAL_W}x{NORMAL_H}+{x}+{y}")
    root.update_idletasks()  # 讓 Tk 真的把這個尺寸記成「正常尺寸」
    try:
        root.state("zoomed")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
