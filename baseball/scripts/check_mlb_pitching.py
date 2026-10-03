"""MLB投手の投球回・防御率・K%・BB%が、MLB公式（statsapi.mlb.com）のシーズン成績と合っているかを調べる（読み取り専用）。

  python baseball/scripts/check_mlb_pitching.py --year 2026 2025

docs/baseball/data/MLB/{年}年/公式戦/pitcher_cards_numeric/index.json を、公式の「レギュラーシーズン」成績と
選手名（アクセント・記号を除いた英数字）＋チームで突き合わせ、指標ごとのずれを表示する。
"""
from __future__ import annotations

import argparse
import json
import re
import unicodedata
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[2]


def norm(name: str) -> str:
    s = unicodedata.normalize("NFKD", name or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_outs(ip) -> int | None:
    """'33.1' / 33.1 / '33.2' → アウト数（小数部は0〜2のアウト数）。"""
    try:
        a, _, b = str(ip).partition(".")
        return int(a) * 3 + int(b or 0)
    except ValueError:
        return None


def official(year: str) -> dict:
    url = ("https://statsapi.mlb.com/api/v1/stats?stats=season&group=pitching&gameType=R"
           f"&season={year}&playerPool=ALL&sportId=1&limit=3000")
    data = requests.get(url, timeout=60).json()
    out: dict = {}
    for sp in data["stats"][0]["splits"]:
        out.setdefault(norm(sp["player"]["fullName"]), []).append(sp)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", nargs="+", default=["2026"])
    ap.add_argument("--min-outs", type=int, default=30, help="これ未満のアウト数（約10回未満）の投手は除く")
    args = ap.parse_args()
    for year in args.year:
        idx = json.loads((ROOT / "docs/baseball/data/MLB" / f"{year}年" / "公式戦" / "pitcher_cards_numeric" / "index.json").read_text(encoding="utf-8"))["players"]
        off = official(year)
        rows, unmatched = [], 0
        for p in idx:
            outs = to_outs(p.get("innings"))
            if outs is None or outs < args.min_outs:
                continue
            cands = off.get(norm(p["name"]), [])
            sp = None
            if len(cands) == 1:
                sp = cands[0]
            elif len(cands) > 1:   # 同姓同名は投球回が最も近い方
                sp = min(cands, key=lambda c: abs((to_outs(c["stat"].get("inningsPitched")) or 0) - outs))
            if not sp:
                unmatched += 1
                continue
            rows.append((p, sp["stat"], outs))
        print(f"\n===== {year}年: 照合 {len(rows)}人 / 公式に見つからない {unmatched}人 =====")
        checks = []
        for p, st, outs in rows:
            bf = num(st.get("battersFaced"))
            checks.append(("投球回(アウト数)", p, outs, to_outs(st.get("inningsPitched")), 0))
            checks.append(("防御率", p, num(p.get("era")), num(st.get("era")), 0.011))
            if bf:
                checks.append(("K%", p, num(p.get("k_pct")), round(num(st.get("strikeOuts")) / bf * 100, 1), 0.11))
                checks.append(("BB%", p, num(p.get("bb_pct")), round(num(st.get("baseOnBalls")) / bf * 100, 1), 0.11))
        for label in ("投球回(アウト数)", "防御率", "K%", "BB%"):
            sel = [c for c in checks if c[0] == label and c[2] is not None and c[3] is not None]
            bad = [c for c in sel if abs(c[2] - c[3]) > c[4] + 1e-9]
            print(f"[{label}] ずれ: {len(bad)}/{len(sel)}人")
            for _, p, a, b, _tol in sorted(bad, key=lambda x: -abs(x[2] - x[3]))[:6]:
                print(f"     {p['name']:<24} {p.get('team')}  当サイト {a} / 公式 {b}  (差 {a - b:+.2f})")


if __name__ == "__main__":
    main()
