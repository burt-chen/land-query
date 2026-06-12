from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from manual_video_common import build_tk_video, capture_tk_window, parse_args


PROJECT = Path(__file__).resolve().parent
TITLE = "國土圖台批次查詢工具"
PURPOSE = "Excel 條件批次查詢地籍資料並匯出"
SCENE_PLAN = [
    ("01", "title", 4.0, None),
    ("02", "files", 6.0, 0),
    ("03", "run", 6.0, 1),
    ("04", "columns", 7.0, 1),
    ("05", "log", 5.0, 2),
]


def build_samples(project_dir: Path) -> None:
    from openpyxl import Workbook

    sample_dir = project_dir / "範例"
    sample_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "查詢條件"
    ws.append(["縣市", "行政區", "大段", "小段", "地號"])
    ws.append(["臺北市", "中正區", "城中段", "", "123"])
    ws.append(["新北市", "板橋區", "府中段", "", "456-1"])
    wb.save(sample_dir / "國土查詢範例.xlsx")


def setup_scene(root, frame, app, scene_id: str, scene_name: str) -> None:
    wrapper = sys.modules[app.__class__.__module__]
    mod = wrapper._tool
    sample_dir = PROJECT / "範例"
    output_path = sample_dir / "國土查詢結果_demo.xlsx"
    input_path = sample_dir / "國土查詢範例.xlsx"
    app._input_path.set(str(input_path))
    app._output_path.set(str(output_path))

    if not getattr(app, "_prepared", None):
        section_index = {
            "臺北市中正區城中": {
                "city": "A", "area": "01", "section": "0001", "office": "AA",
                "縣市": "臺北市", "行政區": "中正區", "事務所": "臺北所", "大段": "城中", "小段": "",
                "所區碼": "AA", "備註": "",
            },
            "新北市板橋區府中": {
                "city": "F", "area": "01", "section": "0002", "office": "FB",
                "縣市": "新北市", "行政區": "板橋區", "事務所": "板橋所", "大段": "府中", "小段": "",
                "所區碼": "FB", "備註": "",
            },
        }
        _, rows = mod._read_xlsx_rows(str(input_path))
        prepared = mod.prepare_input(rows, section_index, app.cfg.get("landno_pattern", mod.DEFAULT_CONFIG["landno_pattern"]))
        app._on_prepared(prepared, section_index)
        app._log("demo Excel 已載入，並完成地段對碼")

    if scene_id != "04":
        for child in find_toplevels(root):
            if child is not root:
                try:
                    child.destroy()
                except Exception:
                    pass

    if scene_id in {"03", "04", "05"} and not app._results_api_done:
        app._results_api_done = [
            {
                "查詢狀態": "成功",
                "縣市": "臺北市",
                "行政區": "中正區",
                "地段": "城中段",
                "地號": "123",
                "面積": "125.4",
                "公告土地現值": "35800",
                "登記日期": "1011018",
                "所有權人": "Demo 所有權人",
                "統一編號": "A123456789",
                "經度": "121.513",
                "緯度": "25.044",
                "查詢縣市": "臺北市",
                "查詢區": "中正區",
                "查詢地段": "城中段",
                "查詢地號": "123",
            }
        ]
        app._results_api_fail = []
        app._refill_done_tree()
        app._refill_fail_tree()
        app._refresh_api_tab_counts()
        app._refresh_api_buttons()
        app._run_status_api.set("demo 查詢完成：成功 1 筆")
        app._log("demo 查詢結果已建立，可匯出 Excel")
        mod.export_results_template(app._results_api_done, str(output_path))
        app._log(f"demo 結果已匯出：{output_path}")

    if scene_id == "04":
        has_dialog = any(child is not root for child in find_toplevels(root))
        if not has_dialog:
            app._open_export_cols_dialog()
        root.update()
        for child in find_toplevels(root):
            if child is not root:
                child.lift()
                child.update()


def find_toplevels(widget):
    found = []
    for child in widget.winfo_children():
        try:
            if child.winfo_toplevel() is child:
                found.append(child)
            found.extend(find_toplevels(child))
        except Exception:
            pass
    return found


def process_land_query_frame(root, frame, app, scene_id: str, scene_name: str, image):
    if scene_id == "04":
        for child in find_toplevels(root):
            if child is not root:
                child.update()
                return capture_tk_window(child)
    return image


if __name__ == "__main__":
    args = parse_args()
    build_tk_video(
        PROJECT,
        TITLE,
        PURPOSE,
        SCENE_PLAN,
        sample_builder=build_samples,
        setup_scene=setup_scene,
        frame_processor=process_land_query_frame,
        silent=args.silent,
        keep_temp=args.keep_temp,
    )
