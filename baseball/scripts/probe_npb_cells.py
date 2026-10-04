"""NPB箱スコアの打席別結果セル（1回〜）の書式を洗い出す読み取り専用スクリプト。
1つのセルに複数の打席（同じイニングに2度打席に立つ）が入る場合の書式を確認し、二塁打・三塁打の数え漏れの有無を調べる。"""
from __future__ import annotations
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/baseball/プロ野球/2026年/1軍/レギュラーシーズン/raw")
cells = Counter()
multi = Counter()
ends = Counter()
inn_cols = Counter()
for f in sorted(root.glob("*/all_games_*.xlsx")):
    try:
        df = pd.read_excel(f, sheet_name="打撃成績")
    except Exception as e:  # noqa: BLE001
        print("skip", f, e)
        continue
    for c in df.columns:
        if re.match(r"^\d+回$", str(c)):
            inn_cols[str(c)] += 1
    for _, r in df.iterrows():
        for c in df.columns:
            if not re.match(r"^\d+回$", str(c)):
                continue
            v = r[c]
            if pd.isna(v) or str(v).strip() in ("", "nan"):
                continue
            s = str(v).strip()
            cells[s] += 1
            if re.search(r"\s|,|、|/|\n|・", s):
                multi[s] += 1
            ends[re.sub(r"^.*?([^\d２３]*)([２３2-3]?)$", r"\2", s)] += 1
print("イニング列:", sorted(inn_cols, key=lambda x: int(x[:-1])))
print("distinct cells:", len(cells))
print("複数打席らしいセル:", multi.most_common(40))
print("末尾（2/3/空）:", ends.most_common(8))
tw = [(s, n) for s, n in cells.items() if re.search(r"[２2３3]", s)]
print("2/3を含むセルの種類:", sorted(tw, key=lambda x: -x[1])[:60])
