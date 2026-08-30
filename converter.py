#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convert 非小酋 (feixiaoqiu) wish history xlsx to Paimon.moe Excel format.

Example:
    python converter.py "非小酋抽卡记录.xlsx" -o paimonmoe_wish_history_converted.xlsx
"""

from __future__ import annotations

import argparse
import datetime
import json
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 非小酋 sheet -> (Paimon.moe sheet, banner category key in data/banners.json)
SHEET_MAP = {
    "角色活动祈愿": ("Character Event", "characters"),
    "武器活动祈愿": ("Weapon Event", "weapons"),
    "常驻祈愿": ("Standard", "standard"),
    "新手祈愿": ("Beginners' Wish", "beginners"),
    # Paimon.moe Excel import currently ignores this extra sheet, but we keep the data.
    "集录祈愿": ("Chronicled Wish", "chronicled"),
}

HEADERS = ["Type", "Name", "Time", "⭐", "Pity", "#Roll", "Group", "Banner", "Part"]


def parse_time(value) -> datetime.datetime:
    if isinstance(value, datetime.datetime):
        return value
    return datetime.datetime.strptime(str(value).strip(), "%Y-%m-%d %H:%M:%S")


def load_data(data_dir: Path):
    with open(data_dir / "zh_to_en_items.json", encoding="utf-8") as f:
        zh2en = json.load(f)
    with open(data_dir / "banners.json", encoding="utf-8") as f:
        bundled = json.load(f)
    return zh2en, bundled["banners"], bundled["dual"]


def find_banner(entries, dt):
    for banner in entries:
        start = parse_time(banner["start"])
        end = parse_time(banner["end"])
        if start <= dt < end:
            return banner
    return None


def banner_id(banner) -> tuple:
    # The same banner name may rerun many times; image+time range makes it unique.
    return (banner["name"], banner["image"], banner["start"], banner["end"])


def get_second_banner(dual, primary):
    if primary is None:
        return None
    key = f"{primary['name']} {primary['image']}"
    pair = dual.get(key)
    if pair and len(pair) >= 2:
        return pair[1]
    return None


def build_group_part(processed, dual, use_heuristic: bool) -> dict:
    """Return {(banner_id, time): 'Wish 2' or ''} for character event rows.

    The source file does not contain gacha_type 301/400, so this is heuristic:
    if a 5-star in the same timestamp equals the second banner's shortName,
    the whole wish action is treated as Wish 2.
    """
    if not use_heuristic:
        return {}

    result = {}
    groups = defaultdict(list)
    for idx, (banner, row) in enumerate(processed):
        groups[(banner_id(banner), row["time"])].append(idx)

    for (bid, _time), idxs in groups.items():
        primary = processed[idxs[0]][0]
        second = get_second_banner(dual, primary)
        if second is None:
            result[(bid, _time)] = ""
            continue

        part = ""
        for idx in idxs:
            row = processed[idx][1]
            if row["star"] == 5:
                row_name = row.get("name") or row.get("en_name", "")
                if row_name == second.get("shortName"):
                    part = "Wish 2"
                    # A ten-pull can theoretically contain both featured 5-stars;
                    # if a later row is Wish 2, prefer Wish 2 for the whole group.
        result[(bid, _time)] = part
    return result


def convert_sheet(source_ws, target_sheet_name, banner_key, banners, dual, use_heuristic: bool, warnings: list):
    rows = []
    for raw in source_ws.iter_rows(min_row=2, values_only=True):
        if not raw or raw[0] is None:
            continue
        time_str, zh_name, zh_type, star, _total, pity, _gacha_id = raw[:7]
        if zh_name is None:
            continue

        en_name = zh2en.get(str(zh_name).strip())
        if en_name is None:
            warnings.append(f"[{target_sheet_name}] 找不到英文名: {zh_name}")
            en_name = str(zh_name)

        if zh_type == "角色":
            en_type = "Character"
        elif zh_type == "武器":
            en_type = "Weapon"
        else:
            en_type = ""

        try:
            dt = parse_time(time_str)
        except Exception:
            warnings.append(f"[{target_sheet_name}] 无法解析时间: {time_str}")
            continue

        rows.append({
            "time": dt,
            "time_str": str(time_str).strip(),
            "name": en_name,
            "type": en_type,
            "star": int(star) if star is not None else 0,
            "pity": int(pity) if pity is not None else 0,
        })

    # Keep source order (非小酋 is normally already chronological).
    rows.sort(key=lambda r: (r["time"], r["time_str"]))

    # Attach the banner occurrence to every row.
    processed = []
    skipped = 0
    for row in rows:
        banner = find_banner(banners.get(banner_key, []), row["time"])
        if banner is None:
            skipped += 1
            warnings.append(f"[{target_sheet_name}] 没有匹配卡池: {row['time_str']} {row['name']}")
            continue
        processed.append((banner, row))
    if skipped:
        warnings.append(f"[{target_sheet_name}] 共跳过 {skipped} 条无法匹配卡池的记录")

    group_part = build_group_part(processed, dual, use_heuristic) if banner_key == "characters" else {}

    output_rows = []
    current_bid = None
    roll = 0
    group = 0
    last_time = None

    for banner, row in processed:
        bid = banner_id(banner)

        if bid != current_bid:
            current_bid = bid
            roll = 0
            group = 0
            last_time = None

        roll += 1
        if row["time"] != last_time:
            group += 1
            last_time = row["time"]

        part = ""
        display_banner = banner["name"]
        if banner_key == "characters":
            second = get_second_banner(dual, banner)
            if second is not None:
                part = group_part.get((bid, row["time"]), "")
                if part == "Wish 2":
                    display_banner = second["name"]

        output_rows.append([
            row["type"],
            row["name"],
            row["time_str"],
            row["star"],
            row["pity"],
            roll,
            group,
            display_banner,
            part,
        ])

    return output_rows


def write_output(output_path: Path, sheet_rows: dict, banners, export_date: str):
    wb = Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True)
    thin = Side(style="thin", color="FFDDDDDD")
    border = Border(left=thin, right=thin, bottom=thin)

    # Wish history sheets
    for sheet_name in [
        "Character Event",
        "Weapon Event",
        "Standard",
        "Beginners' Wish",
        # Extra sheet: Paimon.moe Excel import ignores it, but data is preserved.
        "Chronicled Wish",
    ]:
        ws = wb.create_sheet(sheet_name)
        ws.append(HEADERS)
        for cell in ws[1]:
            cell.font = header_font
        for row in sheet_rows.get(sheet_name, []):
            ws.append(row)
        # simple column widths
        widths = [10, 32, 22, 5, 6, 7, 7, 32, 9]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        for r in ws.iter_rows(min_row=2):
            for cell in r:
                cell.border = border
        ws.freeze_panes = "A2"

    # Banner List
    ws = wb.create_sheet("Banner List")
    ws.append(["Name", "Start", "End"])
    for cell in ws[1]:
        cell.font = header_font
    # Same order as Paimon.moe export: beginners, standard, characters, weapons, chronicled
    for category in ["beginners", "standard", "characters", "weapons", "chronicled"]:
        for b in banners.get(category, []):
            ws.append([b["name"], b["start"], b["end"]])
    for col, width in zip("ABC", [32, 19, 19]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"

    # Information
    ws = wb.create_sheet("Information")
    ws.append(["Paimon.moe Wish History Export"])
    ws.append(["Version", 3])
    ws.append(["Export Date", export_date])
    ws.merge_cells("A1:B1")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28

    wb.save(output_path)


def default_data_dir() -> Path:
    """Return the data directory, including inside a PyInstaller onefile bundle."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "data"


def auto_input_path() -> Path:
    candidates = sorted(Path(".").glob("非小酋抽卡记录*.xlsx"))
    if not candidates:
        raise FileNotFoundError("当前目录没有找到 非小酋抽卡记录*.xlsx")
    return candidates[0]


def convert_paimon_file(
    input_path,
    output_path,
    data_dir=None,
    use_heuristic=True,
    print_progress=False,
):
    """Convert one Feixiaoqiu xlsx to Paimon.moe xlsx. Returns warnings list."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    if data_dir is None:
        data_dir = default_data_dir()
    data_dir = Path(data_dir)

    global zh2en, banners, dual
    zh2en, banners, dual = load_data(data_dir)

    wb = load_workbook(input_path, read_only=True, data_only=True)
    warnings = []
    sheet_rows = {}
    for source_sheet, (target_sheet, banner_key) in SHEET_MAP.items():
        if source_sheet not in wb.sheetnames:
            warnings.append(f"源文件缺少 Sheet: {source_sheet}")
            continue
        if print_progress:
            print(f"转换: {source_sheet} -> {target_sheet}")
        rows = convert_sheet(
            wb[source_sheet],
            target_sheet,
            banner_key,
            banners,
            dual,
            use_heuristic=use_heuristic,
            warnings=warnings,
        )
        sheet_rows[target_sheet] = rows
        if print_progress:
            print(f"  行数: {len(rows)}")

    export_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    write_output(output_path, sheet_rows, banners, export_date)
    return warnings


def main(argv=None):
    parser = argparse.ArgumentParser(description="非小酋抽卡记录 -> Paimon.moe Excel 转换器")
    parser.add_argument("input", nargs="?", help="非小酋导出的 xlsx 文件；缺省自动查找")
    parser.add_argument("-o", "--output", default="paimonmoe_wish_history_converted.xlsx", help="输出 xlsx 路径")
    parser.add_argument("--data-dir", default=None, help="数据目录，默认使用内置 data")
    parser.add_argument("--no-heuristic", action="store_true", help="禁用双卡池 Wish 2 启发式识别")
    args = parser.parse_args(argv)

    data_dir = Path(args.data_dir) if args.data_dir else default_data_dir()
    if not data_dir.is_dir():
        parser.error(f"找不到数据目录: {data_dir}")

    input_path = Path(args.input) if args.input else auto_input_path()
    if not input_path.is_file():
        parser.error(f"输入文件不存在: {input_path}")

    print(f"读取: {input_path}")
    warnings = convert_paimon_file(
        input_path,
        args.output,
        data_dir=data_dir,
        use_heuristic=not args.no_heuristic,
        print_progress=True,
    )

    print(f"\n输出: {args.output}")
    if warnings:
        print("\n警告:")
        for w in warnings:
            print(" -", w)
    else:
        print("无警告")


if __name__ == "__main__":
    main()
