#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Update bundled data files from the public Paimon.moe GitHub repository.

Run this when new Genshin versions / banners / items are released:

    python update_data.py
"""

import json
import re
import urllib.request
from pathlib import Path

DATA_DIR = Path("data")
BASE = "https://raw.githubusercontent.com/MadeBaruna/paimon-moe/main/"


def fetch(path):
    url = BASE + path
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return urllib.request.urlopen(req, timeout=60).read().decode("utf-8")


def parse_banners_js(data):
    data2 = re.sub(r"//[^\n]*", "", data)
    pat = re.compile(r"^\s{2}([A-Za-z_][A-Za-z0-9_-]*)\s*:\s*\[", re.M)
    banners = {}
    for m in pat.finditer(data2):
        key = m.group(1) or m.group(2) or m.group(3)
        i = m.end()
        depth = 1
        while i < len(data2) and depth > 0:
            if data2[i] == "[":
                depth += 1
            elif data2[i] == "]":
                depth -= 1
            i += 1
        block = data2[m.end():i - 1]
        objects = []
        j = 0
        while True:
            j = block.find("{", j)
            if j == -1:
                break
            d = 0
            k = j
            while k < len(block):
                if block[k] == "{":
                    d += 1
                elif block[k] == "}":
                    d -= 1
                if d == 0:
                    objects.append(block[j:k + 1])
                    j = k + 1
                    break
                k += 1
            else:
                break
        entries = []
        for obj in objects:
            def field(name):
                mm = re.search(r"\b" + re.escape(name) + r"\s*:\s*([\"']?)(.*?)\1\s*[,}\n]", obj, re.S)
                return mm.group(2).strip() if mm else None
            name = field("name")
            if name is None:
                continue
            entries.append({
                "name": name,
                "image": field("image"),
                "shortName": field("shortName"),
                "start": field("start"),
                "end": field("end"),
            })
        banners[key] = entries
    return banners


def parse_banners_dual(data):
    pat = re.compile(r'''"([^"]*)"\s*:\s*\[|'([^']*)'\s*:\s*\[''')
    dual = {}
    for m in pat.finditer(data):
        key = m.group(1) or m.group(2) or m.group(3) if m.group(1) is not None else m.group(2)
        i = m.end()
        depth = 1
        while i < len(data) and depth > 0:
            if data[i] == "[":
                depth += 1
            elif data[i] == "]":
                depth -= 1
            i += 1
        block = data[m.end():i - 1]
        objects = []
        j = 0
        while True:
            j = block.find("{", j)
            if j == -1:
                break
            d = 0
            k = j
            while k < len(block):
                if block[k] == "{":
                    d += 1
                elif block[k] == "}":
                    d -= 1
                if d == 0:
                    objects.append(block[j:k + 1])
                    j = k + 1
                    break
                k += 1
            else:
                break
        entries = []
        for obj in objects:
            def field(name):
                mm = re.search(r"\b" + re.escape(name) + r"\s*:\s*([\"']?)(.*?)\1\s*[,}\n]", obj, re.S)
                return mm.group(2).strip() if mm else None
            name = field("name")
            if name is None:
                continue
            entries.append({
                "name": name,
                "image": field("image"),
                "shortName": field("shortName"),
            })
        if entries:
            dual[key] = entries
    return dual


def parse_id_name_objects(data):
    data = re.sub(r"//[^\n]*", "", data)
    pat = re.compile(r"^\s{2}(?:'([^']+)'|\"([^\"]+)\"|([A-Za-z_][A-Za-z0-9_-]*))\s*:\s*\{", re.M)
    result = {}
    for m in pat.finditer(data):
        key = m.group(1) or m.group(2) or m.group(3)
        i = m.end()
        depth = 1
        while i < len(data) and depth > 0:
            if data[i] == "{":
                depth += 1
            elif data[i] == "}":
                depth -= 1
            i += 1
        block = data[m.end():i - 1]

        def field(name):
            mm = re.search(r"\b" + re.escape(name) + r"\s*:\s*([\"'])(.*?)\1", block, re.S)
            return mm.group(2) if mm else None

        idv = field("id")
        namev = field("name")
        if idv and namev:
            result[idv] = namev
    return result


def main():
    DATA_DIR.mkdir(exist_ok=True)

    print("Downloading item names...")
    items = json.loads(fetch("src/locales/items/zh.json"))
    zh2en = {v: k for k, v in items.items()}
    (DATA_DIR / "zh_to_en_items.json").write_text(
        json.dumps(zh2en, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print("Downloading banner data...")
    banners = parse_banners_js(fetch("src/data/banners.js"))
    dual = parse_banners_dual(fetch("src/data/bannersDual.js"))
    (DATA_DIR / "banners.json").write_text(
        json.dumps({"banners": banners, "dual": dual}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("Downloading item id data...")
    chars = parse_id_name_objects(fetch("src/data/characters.js"))
    weaps = parse_id_name_objects(fetch("src/data/weaponList.js"))
    en2id = {}
    for idv, name in {**chars, **weaps}.items():
        en2id[name] = idv
    (DATA_DIR / "item_ids.json").write_text(
        json.dumps(en2id, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print("Downloading UIGF numeric item id data...")
    uigf_req = urllib.request.Request(
        "https://api.uigf.org/dict/genshin/all.json",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    uigf_all = json.loads(urllib.request.urlopen(uigf_req, timeout=60).read().decode("utf-8"))
    numeric_ids = {
        "chs": {k: str(v) for k, v in uigf_all.get("chs", {}).items()},
        "en": {k: str(v) for k, v in uigf_all.get("en", {}).items()},
    }
    (DATA_DIR / "uigf_numeric_item_ids.json").write_text(
        json.dumps(numeric_ids, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print("Done. Data files are up to date.")


if __name__ == "__main__":
    main()
