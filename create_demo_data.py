#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate small synthetic demo xlsx files (no real user data)."""

import datetime
import json
import random
from pathlib import Path

from openpyxl import Workbook

random.seed(20260830)

DATA_DIR = Path("data")
EXAMPLES_DIR = Path("examples")

# Chinese item name -> (English type, star)
ITEMS = {
    # 3-star weapons
    "弹弓": ("武器", 3),
    "魔导绪论": ("武器", 3),
    "黎明神剑": ("武器", 3),
    "黑缨枪": ("武器", 3),
    "翡玉法球": ("武器", 3),
    "沐浴龙血的剑": ("武器", 3),
    "铁影阔剑": ("武器", 3),
    "冷刃": ("武器", 3),
    "鸦羽弓": ("武器", 3),
    "飞天御剑": ("武器", 3),
    # 4-star characters / weapons
    "行秋": ("角色", 4),
    "香菱": ("角色", 4),
    "班尼特": ("角色", 4),
    "菲谢尔": ("角色", 4),
    "砂糖": ("角色", 4),
    "祭礼大剑": ("武器", 4),
    "雨裁": ("武器", 4),
    # 5-star characters
    "莫娜": ("角色", 5),
    "刻晴": ("角色", 5),
    "迪卢克": ("角色", 5),
    "七七": ("角色", 5),
}

SHEET_BANNER_KEY = {
    "角色活动祈愿": "characters",
    "武器活动祈愿": "weapons",
    "常驻祈愿": "standard",
    "集录祈愿": "chronicled",
    "新手祈愿": "beginners",
}

# How many rows per demo sheet (small, enough to show the format)
ROW_COUNTS = {
    "角色活动祈愿": 60,
    "武器活动祈愿": 20,
    "常驻祈愿": 15,
    "集录祈愿": 10,
    "新手祈愿": 0,
}


def parse_dt(s):
    return datetime.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")


def pick_banner(banners, key):
    entries = banners.get(key, [])
    return random.choice(entries)


def random_time(banner):
    start = parse_dt(banner["start"])
    end = parse_dt(banner["end"])
    # avoid very long permanent banner range
    if (end - start).days > 3650:
        start = datetime.datetime(2025, 1, 1, 0, 0, 0)
        end = datetime.datetime(2025, 6, 1, 0, 0, 0)
    span = max(1, int((end - start).total_seconds()))
    return start + datetime.timedelta(seconds=random.randint(0, span - 1))


def main():
    with open(DATA_DIR / "banners.json", encoding="utf-8") as f:
        banners = json.load(f)["banners"]

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, banner_key in SHEET_BANNER_KEY.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(["时间", "名称", "类别", "星级", "总次数", "保底内", "祈愿 Id"])
        total = 0
        pity = 0
        for i in range(ROW_COUNTS[sheet_name]):
            banner = pick_banner(banners, banner_key)
            dt = random_time(banner)
            name = random.choice(list(ITEMS.keys()))
            category, star = ITEMS[name]
            total += 1
            pity += 1
            displayed_pity = pity
            if star == 5:
                pity = 0
            # Fake wish id: timestamp + fake sequence. It is not a real UID/wish id.
            fake_id = f"{int(dt.timestamp() * 1000):013d}{900000 + i:06d}"
            ws.append([
                dt.strftime("%Y-%m-%d %H:%M:%S"),
                name,
                category,
                star,
                total,
                displayed_pity,
                fake_id,
            ])

    EXAMPLES_DIR.mkdir(exist_ok=True)
    out = EXAMPLES_DIR / "input_demo.xlsx"
    wb.save(out)
    print(f"生成: {out}")


if __name__ == "__main__":
    main()
