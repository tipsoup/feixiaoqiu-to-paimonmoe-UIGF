#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convert 非小酋 (feixiaoqiu) wish history xlsx to UIGF JSON.

Default output is the latest UIGF v4.2.
Use --uigf-version v3.0 for Teyvat Guide / 提瓦特指南 compatibility.

Example:
    python to_uigf.py "非小酋抽卡记录.xlsx" --uid 123456789 -o uigf_wish_history.json
"""

from __future__ import annotations

import argparse
import datetime
import json
from pathlib import Path

from openpyxl import load_workbook

from converter import (
    SHEET_MAP,
    banner_id,
    build_group_part,
    default_data_dir,
    find_banner,
    get_second_banner,
    load_data,
    parse_time,
)

# 非小酋 sheet -> (gacha_type, uigf_gacha_type)
UIGF_SHEET_MAP = {
    "角色活动祈愿": ("301", "301"),
    "武器活动祈愿": ("302", "302"),
    "常驻祈愿": ("200", "200"),
    "新手祈愿": ("100", "100"),
    "集录祈愿": ("500", "500"),
}

# Source banner category used for matching time to banner
SHEET_BANNER_KEY = {
    "角色活动祈愿": "characters",
    "武器活动祈愿": "weapons",
    "常驻祈愿": "standard",
    "新手祈愿": "beginners",
    "集录祈愿": "chronicled",
}


def load_item_ids(data_dir: Path):
    with open(data_dir / "item_ids.json", encoding="utf-8") as f:
        return json.load(f)


def load_numeric_item_ids(data_dir: Path):
    with open(data_dir / "uigf_numeric_item_ids.json", encoding="utf-8") as f:
        return json.load(f)


def convert_to_uigf_file(
    input_path,
    output_path,
    uid,
    lang="zh-cn",
    timezone=8,
    data_dir=None,
    use_heuristic=True,
    uigf_version="v4.2",
    print_progress=False,
):
    """Convert one Feixiaoqiu xlsx to UIGF v4.2 JSON. Returns warnings list."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    if data_dir is None:
        data_dir = default_data_dir()
    data_dir = Path(data_dir)

    zh2en, banners, dual = load_data(data_dir)
    en2id = load_item_ids(data_dir)
    numeric_ids = load_numeric_item_ids(data_dir)

    wb = load_workbook(input_path, read_only=True, data_only=True)
    warnings = []
    records = []

    for source_sheet, (gacha_type, uigf_gacha_type) in UIGF_SHEET_MAP.items():
        if source_sheet not in wb.sheetnames:
            warnings.append(f"源文件缺少 Sheet: {source_sheet}")
            continue
        banner_key = SHEET_BANNER_KEY[source_sheet]

        processed = []
        for raw in wb[source_sheet].iter_rows(min_row=2, values_only=True):
            if not raw or raw[0] is None:
                continue
            time_str, zh_name, zh_type, star, _total, _pity, wish_id = raw[:7]
            if zh_name is None:
                continue
            en_name = zh2en.get(str(zh_name).strip())
            if en_name is None:
                warnings.append(f"[{source_sheet}] 找不到英文名: {zh_name}")
                en_name = str(zh_name)

            try:
                dt = parse_time(time_str)
            except Exception:
                warnings.append(f"[{source_sheet}] 无法解析时间: {time_str}")
                continue

            banner = find_banner(banners.get(banner_key, []), dt)
            if banner is None:
                warnings.append(f"[{source_sheet}] 没有匹配卡池: {time_str} {zh_name}")
                continue

            processed.append((banner, {
                "time": dt,
                "time_str": str(time_str).strip(),
                "zh_name": str(zh_name).strip(),
                "en_name": en_name,
                "zh_type": str(zh_type).strip(),
                "star": int(star) if star is not None else 0,
                "wish_id": str(wish_id).strip() if wish_id is not None else "",
            }))

        group_part = build_group_part(processed, dual, use_heuristic=use_heuristic) if banner_key == "characters" else {}

        for banner, row in processed:
            actual_gacha_type = gacha_type
            if banner_key == "characters":
                second = get_second_banner(dual, banner)
                if second is not None:
                    part = group_part.get((banner_id(banner), row["time"]), "")
                    if part == "Wish 2":
                        actual_gacha_type = "400"

            if lang == "zh-cn":
                name = row["zh_name"]
                item_type = "角色" if row["zh_type"] == "角色" else "武器"
                lang_key = "chs"
            else:
                name = row["en_name"]
                item_type = "Character" if row["zh_type"] == "角色" else "Weapon"
                lang_key = "en"

            item_id = numeric_ids.get(lang_key, {}).get(name, "")
            if not item_id:
                warnings.append(f"[{source_sheet}] 找不到 UIGF item_id: {name}")

            records.append({
                "uigf_gacha_type": uigf_gacha_type,
                "gacha_type": actual_gacha_type,
                "item_id": item_id,
                "count": "1",
                "time": row["time_str"],
                "name": name,
                "item_type": item_type,
                "rank_type": str(row["star"]),
                "id": row["wish_id"],
            })

        if print_progress:
            print(f"转换: {source_sheet} -> UIGF ({len(processed)} 条)")

    records.sort(key=lambda r: r["time"])

    now = datetime.datetime.now()
    info = {
        "uid": uid,
        "lang": lang,
        "export_timestamp": int(now.timestamp()),
        "export_app": "feixiaoqiu-to-paimonmoe",
        "export_app_version": "1.0.0",
    }
    if uigf_version == "v3.0":
        data = {
            "info": {
                **info,
                "uigf_version": "v3.0",
                "export_time": now.strftime("%Y-%m-%d %H:%M:%S"),
                "region_time_zone": timezone,
            },
            "list": records,
        }
    else:
        data = {
            "info": {
                "export_timestamp": info["export_timestamp"],
                "export_app": info["export_app"],
                "export_app_version": info["export_app_version"],
                "version": "v4.2",
            },
            "hk4e": [
                {
                    "uid": uid,
                    "timezone": timezone,
                    "lang": lang,
                    "list": records,
                }
            ],
        }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return warnings


def main(argv=None):
    parser = argparse.ArgumentParser(description="非小酋抽卡记录 -> UIGF v4.2 JSON 转换器")
    parser.add_argument("input", nargs="?", help="非小酋导出的 xlsx 文件；缺省自动查找")
    parser.add_argument("--uid", help="游戏 UID（非小酋文件里没有，需要手动提供）")
    parser.add_argument("--lang", default="zh-cn", choices=["zh-cn", "en"], help="输出语言，默认 zh-cn")
    parser.add_argument("--timezone", type=int, default=8, help="时区偏移，默认 8（Asia）")
    parser.add_argument("-o", "--output", default="uigf_wish_history.json", help="输出 JSON 路径")
    parser.add_argument("--data-dir", default=None, help="数据目录，默认使用内置 data")
    parser.add_argument("--no-heuristic", action="store_true", help="禁用双卡池 400 启发式识别")
    parser.add_argument("--uigf-version", choices=["v3.0", "v4.2"], default="v4.2", help="UIGF 版本，默认 v4.2（最新版）")
    args = parser.parse_args(argv)

    data_dir = Path(args.data_dir) if args.data_dir else default_data_dir()
    if not data_dir.is_dir():
        parser.error(f"找不到数据目录: {data_dir}")

    input_path = Path(args.input) if args.input else None
    if input_path is None:
        candidates = sorted(Path(".").glob("非小酋抽卡记录*.xlsx"))
        if not candidates:
            parser.error("当前目录没有找到 非小酋抽卡记录*.xlsx")
        input_path = candidates[0]
    if not input_path.is_file():
        parser.error(f"输入文件不存在: {input_path}")

    uid = args.uid
    if not uid:
        uid = input("请输入 UID: ").strip()
    if not uid:
        parser.error("UID 不能为空")

    print(f"读取: {input_path}")
    warnings = convert_to_uigf_file(
        input_path,
        args.output,
        uid,
        lang=args.lang,
        timezone=args.timezone,
        data_dir=data_dir,
        use_heuristic=not args.no_heuristic,
        uigf_version=args.uigf_version,
        print_progress=True,
    )

    print(f"输出: {args.output}")
    if warnings:
        print("警告:")
        for w in warnings:
            print(" -", w)
    else:
        print("无警告")


if __name__ == "__main__":
    main()
