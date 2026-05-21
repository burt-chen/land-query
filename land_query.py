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
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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
INPUT_OUTPUT_COLUMNS = ["輸入縣市", "輸入行政區", "輸入大段", "輸入小段", "輸入地號"]


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


def download_section_table(url: str) -> "pandas.DataFrame":
    """下載地段代碼表，整理成 key/city/area/section 四欄。"""
    import urllib3
    import requests
    import pandas as pd

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    resp = requests.get(url, verify=False, timeout=60)
    resp.raise_for_status()
    raw = pd.read_excel(io.BytesIO(resp.content))

    A = raw.iloc[:, 0].fillna("").astype(str).str.strip()
    B = raw.iloc[:, 1].fillna("").astype(str).str.strip()
    C = raw.iloc[:, 2].fillna("").astype(str).str.strip()
    E = raw.iloc[:, 4].fillna("").astype(str).str.strip()
    F = raw.iloc[:, 5].fillna("").astype(str).str.strip()
    G = raw.iloc[:, 6].fillna("").astype(str).str.strip()

    out = pd.DataFrame()
    out["key"] = F + G + A + B
    out["city"] = E.str[0]
    # area = 縣市碼 + 鄉鎮市區代碼（如 E20，給 Selenium 版的 land_area_office 下拉用）
    out["area"] = E.str[0] + E.str[-2:]
    # office = 「所區碼」前 2 碼（地政事務所代碼如 EF，給 API 版的 qryTileMapIndex 用）
    out["office"] = E.str[:2]
    out["section"] = (
        pd.to_numeric(C, errors="coerce").fillna(0).astype(int).astype(str).str.zfill(4)
    )
    return out


def _normalize_input(df: "pandas.DataFrame") -> "pandas.DataFrame":
    df = df.copy()
    df["大段"] = (
        df["大段"].fillna("").astype(str).str.strip().str.replace(r"段$", "", regex=True)
    )
    df["小段"] = (
        df["小段"].fillna("").astype(str).str.strip().str.replace(r"小段$", "", regex=True)
    )
    return df


def _safe_str(v) -> str:
    """把 pandas/numpy 的 NaN、None 轉成空字串；其餘 str() 後 strip。"""
    import pandas as pd
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    # pandas 把 NaN/NaT 轉成字串就是 "nan" / "NaT"，視為空
    if s in ("nan", "NaN", "NaT", "None"):
        return ""
    return s


def prepare_input(
    input_df: "pandas.DataFrame",
    section_df: "pandas.DataFrame",
    landno_pattern: str,
) -> PreparedInput:
    """讀使用者輸入，套地段代碼，過濾找不到代碼 / 地號格式錯誤的資料。"""
    missing = [c for c in INPUT_COLUMNS if c not in input_df.columns]
    if missing:
        raise ValueError(f"input.xlsx 缺少必要欄位: {', '.join(missing)}")

    df = input_df.copy()
    for src, dst in zip(INPUT_COLUMNS, INPUT_OUTPUT_COLUMNS):
        df[dst] = df[src]

    df = _normalize_input(df)
    df["key"] = (
        df["縣市"].astype(str).str.strip()
        + df["行政區"].astype(str).str.strip()
        + df["大段"].astype(str).str.strip()
        + df["小段"].fillna("").astype(str).str.strip()
    )

    merged = df.merge(
        section_df[["key", "city", "area", "office", "section"]], on="key", how="left"
    )

    pattern = re.compile(landno_pattern)
    result = PreparedInput()

    for _, row in merged.iterrows():
        landno = _safe_str(row.get("地號"))
        base_info = {
            "輸入縣市": _safe_str(row["輸入縣市"]),
            "輸入行政區": _safe_str(row["輸入行政區"]),
            "輸入大段": _safe_str(row["輸入大段"]),
            "輸入小段": _safe_str(row["輸入小段"]),
            "輸入地號": _safe_str(row["輸入地號"]),
            "landno": landno,
        }

        # 找不到代碼
        city = _safe_str(row.get("city"))
        if not city:
            result.no_code.append(base_info)
            continue

        # 地號格式
        if not pattern.match(landno):
            result.bad_landno.append(base_info)
            continue

        result.valid.append(
            PreparedRow(
                city=city,
                area=_safe_str(row["area"]),
                section=_safe_str(row["section"]).zfill(4),
                landno=landno,
                輸入縣市=base_info["輸入縣市"],
                輸入行政區=base_info["輸入行政區"],
                輸入大段=base_info["輸入大段"],
                輸入小段=base_info["輸入小段"],
                輸入地號=base_info["輸入地號"],
                office=_safe_str(row.get("office")),
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
    "AA11": "公告土地現值起日",
    "AA12": "公告土地現值",
    "AA16": "申報地價",
    "AA17": "公告土地現值(元/㎡)",
    "AA21": "都市計畫面積",
    "AA22": "非都市計畫面積",
    "AA23": "使用分區",
    "AA24": "登記面積",
    "AA27": "公告日期",
}


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


EXPORT_COLUMNS_TEMPLATE = [
    {"name": "輸入縣市",          "source": "輸入縣市"},
    {"name": "輸入行政區",        "source": "輸入行政區"},
    {"name": "輸入大段",          "source": "輸入大段"},
    {"name": "輸入小段",          "source": "輸入小段"},
    {"name": "輸入地號",          "source": "輸入地號"},
    {"name": "面積",              "source": "面積(平方公尺)"},
    {"name": "使用分區",          "source": None},
    {"name": "使用地類別",        "source": None},
    {"name": "登記日期",          "source": "登記日期",       "transform": "minguo_date"},
    {"name": "公告土地現值",      "source": "申報地價",       "transform": "yuan_per_sqm"},
    {"name": "權利人類別",        "source": None},
    {"name": "地籍連結",          "source": "地籍連結(JSONP)"},
    {"name": "行政區",            "source": "行政區"},
    {"name": "經緯度(度)",        "source": "經緯度(JSONP)"},
    {"name": "經緯度(度分秒)",    "source": "經緯度(度分秒)"},
    {"name": "TWD97",             "source": "TWD97"},  # pyproj 從 cx,cy 換算，格式 "E:xxx N:xxx"
    {"name": "地號",              "source": "地號(JSONP組合)"},
    {"name": "登記日期_1",        "source": None},
    {"name": "登記原因",          "source": None},
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
    import pandas as pd
    cols = list(columns) if columns is not None else list(EXPORT_COLUMNS_TEMPLATE)

    rows_out = []
    for r in results:
        new_row = {}
        for spec in cols:
            name = spec["name"]
            src = spec.get("source")
            tx = spec.get("transform")
            if src is None:
                new_row[name] = ""
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
            new_row[name] = val
        rows_out.append(new_row)

    df = pd.DataFrame(rows_out, columns=[c["name"] for c in cols]).fillna("")
    df.to_excel(path, index=False)


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


def _wgs84_to_twd97(lon: float, lat: float) -> tuple[float, float] | None:
    """WGS84 (EPSG:4326) → TWD97 (EPSG:3826)。pyproj 沒裝就回 None。"""
    try:
        from pyproj import Transformer
    except ImportError:
        return None
    try:
        global _TWD97_TRANSFORMER
        if _TWD97_TRANSFORMER is None:
            _TWD97_TRANSFORMER = Transformer.from_crs("EPSG:4326", "EPSG:3826", always_xy=True)
        e, n = _TWD97_TRANSFORMER.transform(lon, lat)
        return e, n
    except Exception:
        return None


def _api_format_land_record(
    row: "PreparedRow",
    payload: dict,
    location_text: str | None = None,
    tile_index: dict | None = None,
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
        col = RALID_FIELD_MAP.get(k, k)
        data[col] = v

    # lcdetype（土地使用分類比例）— 全攤平，給技術讀者用
    lcdetype = payload.get("lcdetype") or {}
    for k, v in lcdetype.items():
        data[f"lcdetype.{k}"] = v

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

    # TWD97 座標（從 tile_index 的 cx,cy 換算）
    if tile_index and "cx" in tile_index and "cy" in tile_index:
        cx, cy = tile_index["cx"], tile_index["cy"]
        # 經緯度(度) 若 LocationQuery 沒給就從 cx,cy 直接組
        if not data.get("經緯度(度)"):
            data["經緯度(度)"] = f"{cx},{cy}"
        twd97 = _wgs84_to_twd97(cx, cy)
        if twd97:
            e, n = twd97
            data["TWD97_E"] = f"{e:.2f}"
            data["TWD97_N"] = f"{n:.2f}"
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

    return data


# Lazy-init 模組級 transformer，避免每筆都建立一次（昂貴）
_TWD97_TRANSFORMER = None


def run_api_query(
    rows: list[PreparedRow],
    cfg: dict,
    log: Callable[[str], None],
    progress: Callable[[int, int], None],
    should_stop: Callable[[], bool],
    on_row: Callable[[dict], None] | None = None,
) -> list[dict]:
    """純 API 查詢，每筆會打 3 顆 API：
      1. getLandInfoSect — 土地基本資訊 + 所有人 + 公有土地
      2. qryTileMapIndex — 地塊中心經緯度（給 LocationQuery 用、pyproj 換 TWD97）
      3. LocationQuery — 行政區 + 經緯度(度/度分秒) + 國土利用
      + pyproj 把 WGS84 換成 TWD97
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

    def get_tile_index(office: str, sect: str, landno8: str) -> dict | None:
        try:
            r = session.get(
                tile_url,
                params={
                    "type": "2", "flag": "2",
                    "office": office, "sect": sect, "landno": landno8,
                    "alpah": "0.5f",
                },
                timeout=timeout, verify=False,
            )
            if r.status_code != 200:
                return None
            arr = r.json()
            if isinstance(arr, list) and arr:
                return arr[0]
            if isinstance(arr, dict):
                return arr
        except Exception as e:
            log(f"  qryTileMapIndex 失敗: {type(e).__name__}: {e}")
        return None

    def get_location_query(cx: float, cy: float, max_retry: int = 2) -> str | None:
        # 注意：LocationQuery 不能跟主 session 共用！
        # NLSC 後端怪規則：同一個 HTTP session 只回第一次完整資料，之後一律空白。
        # 改用一次性 request；空字串回應視為失敗、重試最多 max_retry 次。
        headers = {
            "Referer": cfg.get("api_referer", DEFAULT_CONFIG["api_referer"]),
            "Origin": "https://maps.nlsc.gov.tw",
            "User-Agent": session.headers.get("User-Agent", "Mozilla/5.0"),
            "X-Requested-With": "XMLHttpRequest",
        }
        for attempt in range(max_retry + 1):
            try:
                r = requests.post(
                    loc_url, data={"center": f"{cx},{cy}"},
                    headers=headers, timeout=timeout, verify=False,
                )
                r.encoding = "utf-8"
                if r.status_code == 200 and r.text.strip():
                    return r.text
                # 空字串：等一下再試
                if attempt < max_retry:
                    time.sleep(0.5 + attempt * 0.5)
            except Exception as e:
                if attempt == max_retry:
                    log(f"  LocationQuery 失敗 ({type(e).__name__}): {e}")
                    return None
                time.sleep(0.5 + attempt * 0.5)
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
            }
            all_results.append(data)
            if on_row:
                try: on_row(data)
                except Exception: pass
            continue

        log(f"查詢第 {idx + 1}/{total} 筆: {row.輸入縣市} {row.輸入行政區} {row.輸入大段}{row.輸入小段} {row.landno}  (API)")

        # --- 1. 主資料 ---
        try:
            r = session.post(
                land_url,
                data={"city": row.city, "sect": row.section, "landno": landno8},
                timeout=timeout, verify=False,
            )
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}")
            payload = r.json()
        except Exception as e:
            data = {
                "輸入縣市": row.輸入縣市, "輸入行政區": row.輸入行政區,
                "輸入大段": row.輸入大段, "輸入小段": row.輸入小段, "輸入地號": row.輸入地號,
                "查詢縣市": row.city, "查詢區": row.area,
                "查詢地段": row.section, "查詢地號": row.landno,
                "查詢狀態": f"getLandInfoSect 失敗: {type(e).__name__}: {e}",
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
        tile_index = get_tile_index(office_code, row.section, landno8)

        # --- 3. LocationQuery 拿行政區/經緯度 ---
        location_text = None
        if tile_index and "cx" in tile_index and "cy" in tile_index:
            location_text = get_location_query(tile_index["cx"], tile_index["cy"])

        # --- 攤平 ---
        data = _api_format_land_record(row, payload, location_text, tile_index)

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
        self._running = False
        self._stop_flag = False
        self._worker: threading.Thread | None = None
        # API 查詢狀態
        self._running_api = False
        self._stop_flag_api = False
        self._worker_api: threading.Thread | None = None
        self._results_api_done: list[dict] = []
        self._results_api_fail: list[dict] = []

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

        return page

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

    # done tree 欄位 = 「#」+ EXPORT_COLUMNS_TEMPLATE 31 欄（跟匯出格式完全一致）
    @property
    def _DONE_COLS(self) -> list[str]:
        return ["#"] + [s["name"] for s in EXPORT_COLUMNS_TEMPLATE]

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
        "登記日期": 150, "公告土地現值": 150, "權利人類別": 100,
        "地籍連結": 320, "行政區": 200,
        "經緯度(度)": 180, "經緯度(度分秒)": 200, "TWD97": 130,
        "地號": 280, "登記日期_1": 110, "登記原因": 90,
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
            "下方表格欄位跟匯出 Excel 完全一致。"
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
        self._run_status_api = tk.StringVar(value="待命")
        ttk.Label(bar, textvariable=self._run_status_api).pack(side="left", padx=12)
        # 匯出按鈕推到最右
        self.btn_export_api = ttk.Button(bar, text="匯出 Excel", command=self._do_export_api, state="disabled")
        self.btn_export_api.pack(side="right")

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
        return frame, tree

    def _apply_api_column_widths(self, tree: ttk.Treeview) -> None:
        """套用欄寬。"""
        for c in list(tree["columns"]):
            tree.heading(c, text=c)
            w = self._API_COL_WIDTHS.get(c, self._DEFAULT_COL_WIDTH)
            anchor = "e" if c == "#" else "w"
            tree.column(c, width=w, anchor=anchor, stretch=False)

    @staticmethod
    def _is_success(data: dict) -> bool:
        status = str(data.get("查詢狀態", "")).strip()
        return status == "" or status == "成功"

    @staticmethod
    def _transform_for_template(data: dict) -> dict:
        """套用 EXPORT_COLUMNS_TEMPLATE 的 transform 規則，得到顯示/匯出用的 dict。"""
        out = {}
        for spec in EXPORT_COLUMNS_TEMPLATE:
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
        tree.insert("", "end", values=values)

    def _insert_fail_row(self, tree: ttk.Treeview, data: dict, idx: int) -> None:
        cols = list(tree["columns"])
        values = []
        for c in cols:
            if c == "#":
                values.append(idx)
            else:
                v = data.get(c, "")
                values.append("" if v is None else str(v))
        tree.insert("", "end", values=values, tags=("row",))

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
        self.btn_start_api.configure(state="disabled")
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
            export_results_template(self._results_api_done, path)
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
        self.btn_stop_api.configure(state="disabled")
        n_done = len(self._results_api_done)
        n_fail = len(self._results_api_fail)
        self._run_status_api.set(f"完成：成功 {n_done} 筆 / 問題 {n_fail} 筆")
        self._log(f"[API] 完成：成功 {n_done} 筆、有問題 {n_fail} 筆")
        self._refresh_api_buttons()

    def _on_run_stopped_api(self) -> None:
        self._running_api = False
        self.btn_start_api.configure(state="normal")
        self.btn_stop_api.configure(state="disabled")
        n_done = len(self._results_api_done)
        n_fail = len(self._results_api_fail)
        self._run_status_api.set(f"已停止（成功 {n_done} 筆 / 問題 {n_fail} 筆）")
        self._log("[API] 查詢已停止")
        self._refresh_api_buttons()

    def _on_run_failed_api(self, msg: str) -> None:
        self._running_api = False
        self.btn_start_api.configure(state="normal")
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
            import pandas as pd
            df = pd.DataFrame(SAMPLE_INPUT_ROWS, columns=INPUT_COLUMNS)
            df.to_excel(path, index=False)
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
                import pandas as pd
                section_df = download_section_table(self.cfg["section_url"])
                input_df = pd.read_excel(path)
                prepared = prepare_input(
                    input_df, section_df, self.cfg.get("landno_pattern", DEFAULT_CONFIG["landno_pattern"]))
                self.root.after(0, lambda: self._on_prepared(prepared))
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._on_prepare_failed(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_prepared(self, prepared: PreparedInput) -> None:
        self._prepared = prepared
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
