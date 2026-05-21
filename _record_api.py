#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性腳本：把 input.xlsx 每筆資料打 3 顆 API，把原始回應 dump 成 markdown / JSON。

用法：
    py _record_api.py

輸出：
    API_完整紀錄.md     ：可讀的 markdown 報告
    API_完整紀錄.json   ：所有原始資料（含 raw response）
"""
from __future__ import annotations

import json
import time
import traceback
from pathlib import Path

import requests
import urllib3

from land_query import (
    DEFAULT_CONFIG,
    _read_xlsx_rows,
    download_section_table,
    prepare_input,
)


def main() -> None:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    cfg_path = Path("config.json")
    cfg = dict(DEFAULT_CONFIG)
    if cfg_path.exists():
        try:
            cfg.update(json.loads(cfg_path.read_text(encoding="utf-8")))
        except Exception:
            pass

    print("== 1. 讀 input.xlsx ==")
    _, input_rows = _read_xlsx_rows("input.xlsx")
    print(f"   {len(input_rows)} 筆")

    print("== 2. 下載地段代碼表 ==")
    section_index = download_section_table(cfg["section_url"])
    print(f"   {len(section_index)} 個段碼")

    print("== 3. 對碼 ==")
    prepared = prepare_input(input_rows, section_index, cfg["landno_pattern"])
    print(
        f"   valid={len(prepared.valid)} "
        f"no_code={len(prepared.no_code)} "
        f"bad={len(prepared.bad_landno)}"
    )

    timeout = float(cfg.get("api_request_timeout", 20))
    delay = float(cfg.get("api_request_delay", 0.5))

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

    records: list[dict] = []

    print("== 4. 開始打 API ==")
    for idx, row in enumerate(prepared.valid, start=1):
        print(f"   [{idx}/{len(prepared.valid)}] {row.輸入縣市} {row.輸入行政區} "
              f"{row.輸入大段}{row.輸入小段} {row.landno}")

        # landno → 8 碼
        try:
            if "-" in row.landno:
                main_n, sub_n = row.landno.split("-", 1)
            else:
                main_n, sub_n = row.landno, "0"
            landno8 = f"{int(main_n):04d}{int(sub_n):04d}"
        except Exception as e:
            print(f"      ! 地號格式異常: {e}")
            continue

        rec: dict = {
            "input": {
                "輸入縣市": row.輸入縣市,
                "輸入行政區": row.輸入行政區,
                "輸入大段": row.輸入大段,
                "輸入小段": row.輸入小段,
                "輸入地號": row.輸入地號,
            },
            "對碼結果": {
                "city": row.city,
                "area": row.area,
                "section": row.section,
                "office": row.office,
                "landno8": landno8,
            },
            "api_calls": {},
        }

        # --- API 1: getLandInfoSect ---
        try:
            r = session.post(
                cfg["api_land_info_url"],
                data={"city": row.city, "sect": row.section, "landno": landno8},
                timeout=timeout, verify=False,
            )
            rec["api_calls"]["getLandInfoSect"] = {
                "request": {
                    "url": cfg["api_land_info_url"],
                    "method": "POST",
                    "data": {"city": row.city, "sect": row.section, "landno": landno8},
                },
                "status_code": r.status_code,
                "response": r.json() if r.status_code == 200 else r.text,
            }
        except Exception as e:
            rec["api_calls"]["getLandInfoSect"] = {
                "error": f"{type(e).__name__}: {e}",
                "trace": traceback.format_exc(),
            }

        # --- API 2: qryTileMapIndex ---
        try:
            params = {
                "type": "2", "flag": "2",
                "office": row.office or row.area,
                "sect": row.section,
                "landno": landno8,
                "alpah": "0.5f",
            }
            r = session.get(cfg["api_tile_index_url"], params=params,
                            timeout=timeout, verify=False)
            tile_resp = r.json() if r.status_code == 200 else r.text
            rec["api_calls"]["qryTileMapIndex"] = {
                "request": {
                    "url": cfg["api_tile_index_url"],
                    "method": "GET",
                    "params": params,
                },
                "status_code": r.status_code,
                "response": tile_resp,
            }
            # 抽 cx,cy 給下一支
            cx = cy = None
            if isinstance(tile_resp, list) and tile_resp:
                cx = tile_resp[0].get("cx"); cy = tile_resp[0].get("cy")
            elif isinstance(tile_resp, dict):
                cx = tile_resp.get("cx"); cy = tile_resp.get("cy")
        except Exception as e:
            rec["api_calls"]["qryTileMapIndex"] = {
                "error": f"{type(e).__name__}: {e}",
                "trace": traceback.format_exc(),
            }
            cx = cy = None

        # --- API 3: LocationQuery（要用新 request 不能共用 session）---
        if cx is not None and cy is not None:
            headers = {
                "Referer": cfg.get("api_referer", DEFAULT_CONFIG["api_referer"]),
                "Origin": "https://maps.nlsc.gov.tw",
                "User-Agent": session.headers["User-Agent"],
                "X-Requested-With": "XMLHttpRequest",
            }
            try:
                r = requests.post(
                    cfg["api_location_query_url"],
                    data={"center": f"{cx},{cy}"},
                    headers=headers, timeout=timeout, verify=False,
                )
                r.encoding = "utf-8"
                rec["api_calls"]["LocationQuery"] = {
                    "request": {
                        "url": cfg["api_location_query_url"],
                        "method": "POST",
                        "data": {"center": f"{cx},{cy}"},
                    },
                    "status_code": r.status_code,
                    "response_text": r.text,
                }
            except Exception as e:
                rec["api_calls"]["LocationQuery"] = {
                    "error": f"{type(e).__name__}: {e}",
                    "trace": traceback.format_exc(),
                }
        else:
            rec["api_calls"]["LocationQuery"] = {"skipped": "no cx,cy from tile index"}

        records.append(rec)
        time.sleep(delay)

    # ---- 寫 JSON ----
    Path("API_完整紀錄.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    print("== 5. 已寫 API_完整紀錄.json ==")

    # ---- 寫 markdown ----
    md = ["# API 完整紀錄", "",
          f"來源:`input.xlsx`,共 {len(records)} 筆有效資料",
          ""]
    for i, rec in enumerate(records, 1):
        inp = rec["input"]; mp = rec["對碼結果"]
        md.append(f"## 第 {i} 筆:{inp['輸入縣市']} {inp['輸入行政區']} "
                  f"{inp['輸入大段']}{inp['輸入小段']} {inp['輸入地號']}")
        md.append("")
        md.append("### 對碼結果")
        md.append("```json")
        md.append(json.dumps(mp, ensure_ascii=False, indent=2))
        md.append("```")
        md.append("")

        for api_name, blob in rec["api_calls"].items():
            md.append(f"### {api_name}")
            md.append("")
            if "request" in blob:
                md.append("**Request**:")
                md.append("```json")
                md.append(json.dumps(blob["request"], ensure_ascii=False, indent=2))
                md.append("```")
                md.append("")
                md.append(f"**HTTP {blob['status_code']} Response**:")
                if "response" in blob:
                    md.append("```json")
                    md.append(json.dumps(blob["response"], ensure_ascii=False, indent=2))
                    md.append("```")
                elif "response_text" in blob:
                    md.append("```")
                    md.append(blob["response_text"])
                    md.append("```")
                md.append("")
            elif "skipped" in blob:
                md.append(f"_skipped: {blob['skipped']}_")
                md.append("")
            else:
                md.append(f"**ERROR**: {blob.get('error')}")
                md.append("")
        md.append("---")
        md.append("")

    Path("API_完整紀錄.md").write_text("\n".join(md), encoding="utf-8")
    print("== 6. 已寫 API_完整紀錄.md ==")


if __name__ == "__main__":
    main()
