# -*- coding: utf-8 -*-
"""Update weekly FX PowerPoint from workbook and chart PNGs.

Current scope
-------------
- Slide 1: Daily heatmap table text + red/blue cell background coloring
- Slide 3: East Asia metric cards
- Slide 4~14: picture/chart placeholders replaced with latest PNGs
- Slide 7: US & Canada metric cards + key-details table
- Slide 1~2 summary text boxes are otherwise left untouched

Usage
-----
python ppt_weekly_updater.py                 # PPT 선택 창 -> 워크북/이미지 폴더 자동 탐색
python ppt_weekly_updater.py --save-as       # 결과를 다른 이름으로 저장 (원본 유지)
python ppt_weekly_updater.py --pptx deck.pptx --workbook FX_Weekly_Automation_Light.xlsx

Lookup rules
------------
--pptx      : argument -> file dialog (weekly deck name changes, so always asked)
--workbook  : argument -> last used (.fx_local.json) -> script folder -> dialog
--assets    : argument -> <workbook folder>/fx_chart_images -> folder dialog
--output    : argument -> (--save-as) save dialog -> overwrite input with backup
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path
from typing import Dict, List

import pandas as pd

from fx_common import (CHART_DIR_NAME, DEFAULT_WORKBOOK, PPTX_TYPES, SCRIPT_DIR, XLSX_TYPES,
                       ask_directory, ask_save_file, remember_path, resolve_file)

IMAGE_MAP: Dict[int, Dict[int, str]] = {
    4: {11: "ma5_usd_krw.png", 10: "ma5_usd_jpy.png", 9: "ma5_usd_cny.png", 12: "ma5_usd_twd.png"},
    5: {7: "ppt_5min_usdkrw_usdjpy.png"},
    6: {9: "bond_swap_spread.png", 10: "swap_basis.png"},
    8: {8: "ma5_dxy.png", 7: "ma5_usd_cad.png"},
    9: {7: "5min_dxy.png"},
    10: {14: "ma5_usd_eur.png", 15: "ma5_usd_gbp.png", 10: "ma5_usd_chf.png", 13: "ma5_usd_sek.png", 11: "ma5_usd_dkk.png", 12: "ma5_usd_rub.png"},
    11: {7: "ppt_5min_usdeur_usdgbp.png"},
    12: {12: "ma5_usd_mxn.png", 10: "ma5_usd_brl.png", 11: "ma5_usd_cop.png", 9: "ma5_usd_ars.png"},
    13: {11: "ma5_usd_inr.png", 13: "ma5_usd_sgd.png", 10: "ma5_usd_idr.png", 14: "ma5_usd_thb.png", 12: "ma5_usd_php.png", 15: "ma5_usd_vnd.png"},
    14: {7: "ma5_usd_aud.png", 8: "ma5_usd_nzd.png"},
}

SLIDE1_HEATMAP_TABLE_SHAPE = 7
SLIDE3_TABLE_MAP = {64: "USD_KRW", 65: "USD_CNY", 71: "USD_JPY", 74: "USD_TWD"}
SLIDE7_SMALL_TABLES = {37: "USD_CAD", 40: "DXY"}
SLIDE7_KEY_DETAILS_SHAPE = 6


def read_sheet(path: Path, sheet: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def load_metrics_by_currency(workbook: Path, sheet_name: str) -> pd.DataFrame:
    df = read_sheet(workbook, sheet_name)
    if df.empty:
        return df
    if "Currency" not in df.columns:
        raise ValueError(f"Sheet {sheet_name} does not contain 'Currency' column.")
    return df.set_index("Currency")


def fmt_current(val: float) -> str:
    if pd.isna(val):
        return "-"
    val = float(val)
    if abs(val) >= 10:
        return f"{val:,.2f}"
    return f"{val:,.4f}"


def fmt_pct(val: float) -> str:
    if pd.isna(val):
        return "-"
    return f"{float(val):.2%}"


def fmt_rsi(val: float) -> str:
    if pd.isna(val):
        return "-"
    return f"{float(val):.1f}"


def fmt_card_values(row: pd.Series) -> List[List[str]]:
    return [
        ["Current", "WoW(%)", "MoM(%)", "YTD(%)"],
        [fmt_current(row.get("Current")), fmt_pct(row.get("WoW(%)")), fmt_pct(row.get("MoM(%)")), fmt_pct(row.get("YTD(%)"))],
        ["Deviation from \\n15 year High (%)", "MDD(%)", "RSI", "Vol 20 Days"],
        [fmt_pct(row.get("Deviation from 15Y High (%)")), fmt_pct(row.get("MDD(%)")), fmt_rsi(row.get("RSI(14)")), fmt_pct(row.get("Vol 20D(%)"))],
    ]


def fmt_small_card_values(row: pd.Series) -> List[List[str]]:
    return [["Current", "WoW(%)"], [fmt_current(row.get("Current")), fmt_pct(row.get("WoW(%)"))]]


def fmt_uscan_key_details(df: pd.DataFrame) -> List[List[str]]:
    dxy = df.loc["DXY"]
    cad = df.loc["USD_CAD"]
    return [
        ["Currency", "DXY", "USD/CAD"],
        ["MoM(%)", fmt_pct(dxy.get("MoM(%)")), fmt_pct(cad.get("MoM(%)"))],
        ["YTD(%)", fmt_pct(dxy.get("YTD(%)")), fmt_pct(cad.get("YTD(%)"))],
        ["Deviation from 15Y High (%)", fmt_pct(dxy.get("Deviation from 15Y High (%)")), fmt_pct(cad.get("Deviation from 15Y High (%)"))],
        ["MDD(%)", fmt_pct(dxy.get("MDD(%)")), fmt_pct(cad.get("MDD(%)"))],
        ["RSI", fmt_rsi(dxy.get("RSI(14)")), fmt_rsi(cad.get("RSI(14)"))],
        ["Vol 20 Days", fmt_pct(dxy.get("Vol 20D(%)")), fmt_pct(cad.get("Vol 20D(%)"))],
    ]


def office_rgb(r: int, g: int, b: int) -> int:
    return int(r) + (int(g) << 8) + (int(b) << 16)


def ordinal_suffix(day: int) -> str:
    if 11 <= day % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def format_heatmap_header(d) -> str:
    d = pd.Timestamp(d)
    return f"{d.strftime('%b')}. {d.day}{ordinal_suffix(d.day)}"


def heatmap_rgb(value, max_abs: float = 0.035):
    """Return a white->red or white->blue gradient color."""
    if pd.isna(value):
        return (255, 255, 255), (0, 0, 0)

    v = float(value)
    strength = min(abs(v) / max_abs, 1.0)

    if v > 0:
        target = (192, 0, 0)
    elif v < 0:
        target = (0, 112, 192)
    else:
        return (255, 255, 255), (0, 0, 0)

    bg = tuple(int(255 + (c - 255) * strength) for c in target)
    font = (255, 255, 255) if strength >= 0.55 else (0, 0, 0)
    return bg, font


def ensure_asset(asset_dir: Path, filename: str) -> Path:
    p = asset_dir / filename
    if not p.exists():
        raise FileNotFoundError(f"Asset not found: {p}")
    return p


def set_table_values(shape, values: List[List[str]]) -> None:
    tbl = shape.Table
    for r, row in enumerate(values, start=1):
        for c, value in enumerate(row, start=1):
            tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = str(value)


def set_cell_text(cell, text: str):
    cell.Shape.TextFrame.TextRange.Text = str(text)


def style_heatmap_cell(cell, value):
    bg, font = heatmap_rgb(value)
    fill = cell.Shape.Fill
    fill.Visible = True
    fill.Solid()
    fill.ForeColor.RGB = office_rgb(*bg)
    fill.Transparency = 0
    cell.Shape.TextFrame.TextRange.Font.Color.RGB = office_rgb(*font)


def replace_shape_with_picture(slide, shape_index: int, image_path: Path):
    shp = slide.Shapes(shape_index)
    left, top, width, height = shp.Left, shp.Top, shp.Width, shp.Height
    try:
        z_order = shp.ZOrderPosition
    except Exception:
        z_order = None
    shp.Delete()
    new_shp = slide.Shapes.AddPicture(str(image_path), LinkToFile=False, SaveWithDocument=True, Left=left, Top=top, Width=width, Height=height)
    if z_order is not None:
        try:
            while new_shp.ZOrderPosition > z_order:
                new_shp.ZOrder(3)
        except Exception:
            pass
    return new_shp


def update_slide1_heatmap(prs, workbook: Path):
    slide = prs.Slides(1)
    shape = slide.Shapes(SLIDE1_HEATMAP_TABLE_SHAPE)
    tbl = shape.Table

    df = read_sheet(workbook, "Daily_Heatmap")
    if df.empty:
        return

    cols = list(df.columns)
    if cols and cols[0] != "Currency":
        cols[0] = "Currency"
        df.columns = cols

    date_cols = list(df.columns[1:6])

    for i in range(5):
        col_idx = 3 + i
        text = ""
        if i < len(date_cols):
            try:
                text = format_heatmap_header(date_cols[i])
            except Exception:
                text = str(date_cols[i])
        set_cell_text(tbl.Cell(1, col_idx), text)

    max_rows = min(len(df), 10)
    for r in range(max_rows):
        row_idx = 2 + r
        row = df.iloc[r]
        set_cell_text(tbl.Cell(row_idx, 2), row.get("Currency", ""))

        for i in range(5):
            col_idx = 3 + i
            value = row[date_cols[i]] if i < len(date_cols) else None
            text = "-" if pd.isna(value) else f"{float(value):.2%}"
            cell = tbl.Cell(row_idx, col_idx)
            set_cell_text(cell, text)
            style_heatmap_cell(cell, value)

    for r in range(max_rows, 10):
        row_idx = 2 + r
        set_cell_text(tbl.Cell(row_idx, 2), "")
        for i in range(5):
            col_idx = 3 + i
            cell = tbl.Cell(row_idx, col_idx)
            set_cell_text(cell, "")
            style_heatmap_cell(cell, None)


def update_slide3_tables(prs, workbook: Path):
    slide = prs.Slides(3)
    metrics = load_metrics_by_currency(workbook, "Metrics_EastAsia")
    for shape_idx, ccy in SLIDE3_TABLE_MAP.items():
        if ccy in metrics.index:
            set_table_values(slide.Shapes(shape_idx), fmt_card_values(metrics.loc[ccy]))


def update_slide7_tables(prs, workbook: Path):
    slide = prs.Slides(7)
    metrics = load_metrics_by_currency(workbook, "Metrics_USCanada")
    for shape_idx, ccy in SLIDE7_SMALL_TABLES.items():
        if ccy in metrics.index:
            set_table_values(slide.Shapes(shape_idx), fmt_small_card_values(metrics.loc[ccy]))
    set_table_values(slide.Shapes(SLIDE7_KEY_DETAILS_SHAPE), fmt_uscan_key_details(metrics))


def replace_all_images(prs, asset_dir: Path):
    for slide_num, shape_map in IMAGE_MAP.items():
        slide = prs.Slides(slide_num)
        for shape_idx in sorted(shape_map.keys(), reverse=True):
            replace_shape_with_picture(slide, shape_idx, ensure_asset(asset_dir, shape_map[shape_idx]))


def update_ppt(pptx: Path, workbook: Path, asset_dir: Path, output: Path | None = None, make_backup: bool = True):
    import win32com.client as win32

    pptx = pptx.resolve()
    workbook = workbook.resolve()
    asset_dir = asset_dir.resolve()
    output = output.resolve() if output else pptx.resolve()

    if not pptx.exists():
        raise FileNotFoundError(pptx)
    if not workbook.exists():
        raise FileNotFoundError(workbook)
    if not asset_dir.exists():
        raise FileNotFoundError(asset_dir)

    if pptx != output:
        shutil.copy2(pptx, output)
    elif make_backup:
        backup = pptx.with_name(pptx.stem + "_before_auto" + pptx.suffix)
        if not backup.exists():
            shutil.copy2(pptx, backup)

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = 1
    pres = app.Presentations.Open(str(output), WithWindow=False)
    try:
        update_slide1_heatmap(pres, workbook)
        update_slide3_tables(pres, workbook)
        update_slide7_tables(pres, workbook)
        replace_all_images(pres, asset_dir)
        pres.Save()
    finally:
        pres.Close()
        app.Quit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", default=None, help="Weekly deck to update; default: file dialog")
    ap.add_argument("--workbook", default=None, help=f"Default: last used -> script folder/{DEFAULT_WORKBOOK} -> dialog")
    ap.add_argument("--assets", default=None, help=f"Chart PNG folder; default: <workbook folder>/{CHART_DIR_NAME}")
    ap.add_argument("--output", default=None, help="Save result here instead of overwriting --pptx")
    ap.add_argument("--save-as", action="store_true", help="Pick the output path in a save dialog")
    ap.add_argument("--ask", action="store_true", help="Always show the workbook dialog too")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    # PPT: weekly file name changes (e.g. 260925_...), so ask every time unless given.
    pptx = resolve_file(args.pptx, key="pptx", title="업데이트할 Weekly FX PPT를 선택하세요",
                        filetypes=PPTX_TYPES, force_dialog=True)

    workbook = resolve_file(args.workbook, key="workbook", title="FX Weekly 워크북(.xlsx)을 선택하세요",
                            filetypes=XLSX_TYPES, default_names=[DEFAULT_WORKBOOK],
                            search_dirs=[SCRIPT_DIR], force_dialog=args.ask)

    if args.assets:
        asset_dir = Path(args.assets).expanduser().resolve()
    else:
        asset_dir = workbook.parent / CHART_DIR_NAME
        if not asset_dir.exists():
            picked = ask_directory(f"차트 이미지 폴더({CHART_DIR_NAME})를 선택하세요", initialdir=workbook.parent)
            if picked is None:
                raise SystemExit("❌ 차트 이미지 폴더 선택 취소됨")
            asset_dir = picked

    output = None
    if args.output:
        output = Path(args.output).expanduser().resolve()
    elif args.save_as:
        output = ask_save_file("결과 PPT 저장 위치", PPTX_TYPES, initialfile=pptx.name, initialdir=pptx.parent)
        if output is None:
            raise SystemExit("❌ 저장 취소됨")

    print(f"PPT      : {pptx}")
    print(f"Workbook : {workbook}")
    print(f"Assets   : {asset_dir}")
    print(f"Output   : {output or pptx}")

    update_ppt(pptx=pptx, workbook=workbook, asset_dir=asset_dir, output=output, make_backup=not args.no_backup)
    if output:
        remember_path("pptx", output)
    print("PPT update complete.")


if __name__ == "__main__":
    main()
