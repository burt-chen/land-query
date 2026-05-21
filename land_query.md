# 國土地籍批次查詢工具

## 功能概要

`land_query.py` 是用於自動批次查詢台灣地籍資料的 GUI 工具（tkinter 分頁介面）。
可獨立執行（`python land_query.py`），也可作為「小工具管理」launcher 的子工具動態載入。

它會：

1. 從內政部地政司 LISP 系統下載地段代碼表（不另存檔案，直接讀進記憶體）
   - 來源：<https://lisp.land.moi.gov.tw/MMS/Handle/DownloadQuerySection.ashx?DownloadType=xls>
2. 讀取 `input.xlsx` 作為查詢條件，並把縣市/行政區/大段/小段轉成 NLSC 查詢用代碼
3. **直接打 NLSC 後端 API（不開瀏覽器）** 逐筆查詢
4. 結果即時顯示在 GUI，並可匯出 Excel

## 用到的 NLSC API

每筆會打 4 顆 API：

| API | 拿什麼 |
|---|---|
| `POST api.nlsc.gov.tw/S09_Ralid/getLandInfoSect` | 土地基本資訊 + 所有人 + 公有土地 |
| `GET landmaps.nlsc.gov.tw/S_Maps/qryTileMapIndex` (JSONP) | 地塊中心經緯度 (cx, cy)、地段中文名、地政事務所代碼 |
| `POST api.nlsc.gov.tw/MapSearch/LocationQuery` | 行政區（含里）、經緯度(度/度分秒)、國土利用現況 |
| `GET api.nlsc.gov.tw/other/GetLandSecInfoNlsc/{city}/{sect}` | 地段元資料（同段共用快取） |

額外用 `pyproj` 把 WGS84 經緯度換算成 TWD97 投影座標。

> ⚠️ **`LocationQuery` 有個 NLSC 後端怪規則**：同一個 HTTP session 只回第一次完整資料，
> 之後一律空白。所以這顆 API 每筆都用獨立 `requests.post()` 打、不共用 session，
> 並且空回應時自動重試最多 2 次。

## GUI 分頁

| 分頁 | 功能 |
|---|---|
| 1. 檔案 | 選輸入檔 `input.xlsx`、預設輸出位置 |
| 2. 預覽 | 下載地段代碼表 + 對碼 + 驗證；分子分頁顯示「可查詢／找不到代碼／地號格式錯誤」 |
| 3. 執行 | 開始 / 停止 / 重試有問題的 / 匯出 Excel / 清空，含 ✓ 完成 / ✗ 有問題 兩個結果分頁 |
| 4. 日誌 | 即時 timestamped 訊息 |
| 5. 設定 | 讀寫 `config.json`：API 端點、逾時、每筆延遲、地號 regex |

### 3. 執行 分頁的兩個結果頁

- **✓ 完成 (N)** — 顯示成功查到資料的筆，欄位 = 匯出 Excel 的 31 欄完全一致（含資料處理：民國日期、元/平方公尺、持分拆分母分子等）
- **✗ 有問題 (N)** — 顯示失敗的筆，只顯示「查詢狀態 + 輸入 5 欄」方便快速看失敗原因（紅底標記）
- **重試有問題的** — 重打「有問題」那邊的查詢；成功的自動搬到「完成」，仍失敗的留在「有問題」
- **匯出 Excel** — 只匯出「完成」那邊的資料

## 設定檔 `config.json`

| 鍵 | 預設 | 說明 |
|---|---|---|
| `section_url` | LISP 地段代碼下載網址 | 改網址用 |
| `landno_pattern` | `^\d{1,4}(-\d{1,4})?$` | 地號合法性 regex |
| `api_land_info_url` | getLandInfoSect | API 端點 |
| `api_sec_info_url` | GetLandSecInfoNlsc | API 端點 |
| `api_tile_index_url` | qryTileMapIndex | API 端點 |
| `api_location_query_url` | LocationQuery | API 端點 |
| `api_referer` | `https://maps.nlsc.gov.tw/` | API 必須帶的 Referer |
| `api_request_timeout` | `20` | 單次 API 等待秒數 |
| `api_request_delay` | `0.5` | 每筆之間延遲（太小會被 NLSC 限流，O 欄度分秒會空白） |

## 匯出 Excel 欄位（31 欄）

完整對照表見 [api_field_mapping.md](api_field_mapping.md)。

簡列：輸入5欄 / 面積 / 使用分區 / 使用地類別 / 登記日期 / 公告土地現值 / 權利人類別 / 地籍連結 / 行政區 / 經緯度(度) / 經緯度(度分秒) / TWD97 / 地號 / 登記日期_1 / 登記原因 / 所有權人 / 統一編號 / 所有權人類別 / 權利範圍類別 / 權利範圍持分_分母 / 權利範圍持分_分子 / 申報地價 / 管理者名稱 / 查詢縣市 / 查詢區 / 查詢地段 / 查詢地號

**資料處理規則**：
- `登記日期` `1011018` → `民國101年10月18日`
- `公告土地現值` `35800` → `35800 元/平方公尺`
- `申報地價` 同上
- `權利範圍持分_分子/分母` 從 `1/3` 拆出 `1` 與 `3`
- 來源 API 沒回的欄位（如使用分區、權利人類別）保持空白

要改欄位設定（增刪、改名、換來源、加處理）：直接改 [land_query.py](land_query.py) 上方的 `EXPORT_COLUMNS_TEMPLATE`。

## 檔案結構

```
國土查詢/
├── land_query.py        # 核心 + GUI 合一（可獨立執行）
├── main_frame.py        # launcher 嵌入殼，提供 create_frame(parent)
├── config.json          # 設定檔
├── requirements.txt     # 依賴套件
├── build.bat            # PyInstaller 打包
├── land_query.md        # 本文件
├── api_field_mapping.md # API 欄位對照表
└── input.xlsx           # 查詢條件範例
```

## 執行方式

### 獨立執行

```powershell
pip install -r requirements.txt
python land_query.py
```

### 透過 launcher

打包後上傳 GitHub Releases，加入 launcher 的 `tools.json`。

### 打包成 exe

```powershell
.\build.bat
```

## 輸入檔 `input.xlsx`

至少含 5 個欄位：`縣市`、`行政區`、`大段`、`小段`、`地號`。

「大段」尾字「段」與「小段」尾字「小段」會被自動去除以利對碼，填寫時可留可省。

## 注意事項

- 下載地段代碼表用 `requests.get(..., verify=False)`，會關閉 SSL 警告
- 「停止」是在當前查詢結束後才中止
- 速度：每筆約 1 秒（含 4 顆 API + 0.5 秒禮貌延遲）；1000 筆約 17 分鐘
- 若 O 欄「經緯度(度分秒)」常空白，去 `config.json` 把 `api_request_delay` 調大

## 相關連結

- 內政部地政司 LISP：<https://lisp.land.moi.gov.tw/>
- 國土測繪中心：<https://maps.nlsc.gov.tw/>
