"""把「影音設備清冊」xlsx 的總表轉成 data/devices.json。

用法：python3 tools/build_devices.py 清冊.xlsx
從 Google 試算表下載：檔案 → 下載 → Microsoft Excel (.xlsx)
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

COLS = ["item", "cat", "loc", "status", "note", "name", "icon", "model", "manual", "ref"]


def clean(v):
    if v is None:
        return ""
    s = str(v)
    s = re.sub(r"<\s*(br|hr)\s*/?>", "\n", s, flags=re.I)
    s = s.replace("✅️", "").replace("✅", "")
    s = re.sub(r"[ 　]{2,}", " ", s)
    return "\n".join(x.strip() for x in s.split("\n") if x.strip())


def main(src):
    ws = openpyxl.load_workbook(src, data_only=True)["總表"]
    out = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        if not any(row):
            continue
        d = {k: clean(v) for k, v in zip(COLS, row)}
        d.pop("icon")
        rec = {"id": f"s{i:02d}"}
        rec.update({k: v for k, v in d.items() if v})
        rec["src"] = "清冊"
        out.append(rec)
    dst = Path(__file__).resolve().parent.parent / "data" / "devices.json"
    dst.parent.mkdir(exist_ok=True)
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"{len(out)} 筆 → {dst}")


if __name__ == "__main__":
    main(sys.argv[1])
