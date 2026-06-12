# 操作手冊影片工作程序

目標：讓每個工具在改版時，都能用一致流程重新產生操作手冊影片，並把影片資訊寫進 `tool_info.json`，讓小工具管理可以同步顯示「操作影片」按鈕。

這份文件是通用流程，不只給 PDFToolkit 使用。不同工具只要替換工具名稱、功能清單、範例素材與產出檔名即可。

## 一、專案內固定結構

建議每個工具專案都維持以下結構：

```text
工具專案/
  make_manual_video.py
  tool_info.json
  manual_video/
    WORKFLOW.md
    script.md
    storyboard.md
    output/
      工具名稱_manual.mp4
  範例/
    demo_file_1
    demo_file_2
    demo_file_3
```

以 PDFToolkit 為例：

```text
PDFToolkit/
  make_manual_video.py
  tool_info.json
  manual_video/
    WORKFLOW.md
    script.md
    storyboard.md
    output/
      PDFToolkit_manual.mp4
  範例/
    demo_pdf_viewer.pdf
    demo_pdf_split.pdf
    demo_replace_text.pdf
```

## 二、第一次建立影片

1. 請 Codex 使用 `operation-manual-video skill` 分析目標工具。
2. 確認工具的主要功能與操作流程。
3. 建立安全的範例資料，放在 `範例/`。
4. 撰寫 `manual_video/script.md`，放字幕或旁白文字。
5. 撰寫 `manual_video/storyboard.md`，放每一段畫面操作流程。
6. 建立 `make_manual_video.py`。
7. 執行 `make_manual_video.py`，產生影片：

```text
manual_video/output/工具名稱_manual.mp4
```

8. 檢查影片：
   - 是否能播放
   - 畫面是否是真實工具畫面
   - 字幕是否沒有擋住重要按鈕
   - 旁白與畫面是否同步
   - 語速是否一致
   - 流程是否能讓使用者照著完成操作

## 三、script.md 建議內容

`script.md` 用來管理影片中的字幕與旁白。建議一段畫面對應一段文字，避免寫成一整篇長文。

```text
# 工具名稱 操作手冊影片旁白

## 01 開場
這是工具名稱的操作手冊影片，接下來會介紹主要功能與基本操作流程。

## 02 功能一
這一段介紹功能一。使用者可以在這裡完成...

## 03 功能二
這一段介紹功能二。建議先...

## 04 結尾
正式處理資料前，建議先確認預覽結果，避免覆蓋原始檔案。
```

## 四、storyboard.md 建議內容

`storyboard.md` 用來描述每一段影片要顯示哪個畫面、要做什麼操作、使用哪一段字幕或旁白。

```text
# 工具名稱 操作手冊影片分鏡

| 段落 | 畫面 | 操作 | 字幕 / 旁白 |
|---|---|---|---|
| 01 | 工具主畫面 | 開啟工具 | script.md / 01 開場 |
| 02 | 功能一畫面 | 載入範例資料並執行預覽 | script.md / 02 功能一 |
| 03 | 功能二畫面 | 設定條件並產生結果 | script.md / 03 功能二 |
| 04 | 結尾畫面 | 顯示完成狀態或注意事項 | script.md / 04 結尾 |
```

## 五、make_manual_video.py 建議責任

`make_manual_video.py` 是重製影片的執行入口，建議固定做以下事情：

1. 建立或確認 `範例/` 測試資料。
2. 啟動真實工具畫面。
3. 依照 `storyboard.md` 操作工具。
4. 擷取真實畫面，不使用示意圖。
5. 加上字幕。
6. 如果需要旁白，依照 `script.md` 分段產生並同步。
7. 輸出到：

```text
manual_video/output/工具名稱_manual.mp4
```

## 六、每次工具改版時判斷影片是否要更新

用這句話判斷：

```text
使用者看舊影片，還能不能順利使用新版工具？
```

判斷規則：

```text
只修 bug / 效能 / 內部邏輯
→ 不一定要重產影片

新增小功能 / 新增選項 / 文字微調
→ 更新 script.md 或 storyboard.md 後重產影片

新增主要功能 / 操作流程改變 / UI 架構改變
→ 請 Codex 重新使用 operation-manual-video skill 規劃影片
```

## 七、改版時的執行流程

```text
1. 修改工具功能
2. 更新版本號
3. 檢查 manual_video/script.md 是否需要修改
4. 檢查 manual_video/storyboard.md 是否需要修改
5. 執行 make_manual_video.py
6. 確認產出的操作手冊影片正常
7. 更新 tool_info.json 的 manual_video 欄位
8. 發 Release 時一併上傳 zip、tool_info.json、manual.mp4
```

## 八、tool_info.json 建議加入欄位

```json
"manual_video": {
  "version": "1.1.0",
  "url": "https://github.com/owner/tool-name/releases/download/v1.1.0/tool-name_manual-v1.1.0.mp4",
  "filename": "tool-name_manual-v1.1.0.mp4",
  "type": "voice",
  "updated_at": "2026-06-10",
  "duration_seconds": 0,
  "sha256": ""
}
```

欄位說明：

```text
version
→ 影片對應的工具版本

url
→ Release 上的影片下載網址

filename
→ Release 上的影片檔名

type
→ silent 或 voice

updated_at
→ 影片最後更新日期

duration_seconds
→ 影片秒數

sha256
→ 影片檔案雜湊，用來確認檔案是否正確
```

如果影片內容跟工具最新版一致，`manual_video.version` 就填工具版本。

如果工具有改版但操作影片不用更新，可以保留舊的 `manual_video.version`。

## 九、打包工具接手的規則

之後 `tools-releases-pack` 可以照這個規則處理：

```text
如果工具專案裡有 make_manual_video.py
→ 詢問是否重新產生操作手冊影片

如果選擇是
→ 執行 make_manual_video.py
→ 找到 manual_video/output/工具名稱_manual.mp4
→ 計算影片大小、sha256、時間
→ 寫入 tool_info.json 的 manual_video 欄位
→ 發 Release 時把影片一起上傳
```

## 十、小工具管理同步規則

`apply_tool_info_gui.py` 只需要同步 `manual_video` 欄位。

它不負責產生影片，只負責把各工具的影片資訊收進 `tools.json`。

整體分工：

```text
各工具專案
→ 產生操作手冊影片

tools-releases-pack
→ 打包並把影片資訊寫進 tool_info.json

apply_tool_info_gui.py
→ 把 manual_video 同步進 tools.json

Launcher
→ 顯示「操作影片」按鈕
```

## 十一、套用到新工具時要填的內容

每個新工具建立影片前，先填好以下資訊：

```text
工具名稱：
啟動方式：
主要功能：
需要的範例資料：
影片是否需要旁白：
影片輸出檔名：
是否需要加入 tool_info.json：
```

填完後，請 Codex 依照本文件建立：

```text
manual_video/script.md
manual_video/storyboard.md
make_manual_video.py
manual_video/output/工具名稱_manual.mp4
```

完成後，操作手冊影片就可以從「手工成果」變成「可重複產生的發版資產」。
