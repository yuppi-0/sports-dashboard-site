"""MLB打者の打率・出塁率・長打率・OPS等が、MLB公式（statsapi.mlb.com）のシーズン成績と合っているかを調べる（読み取り専用）。

  python baseball/scripts/check_mlb_batting.py --year 2026 2025

docs/baseball/data/MLB/{年}年/公式戦/batter_cards_numeric/index.json の値を、公式の「レギュラーシーズン」成績と
選手名（アクセント・記号を除いた英数字）で突き合わせ、指標ごとのずれを表示する。
"""
from __future__ import annotations

import argparse
import json
import re
import unicodedata
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[2]
METRICS = [("avg", "avg", 0.0015), ("obp", "obp", 0.0015), ("slg", "slg", 0.0015), ("ops", "ops", 0.0015),
           ("pa", "plateAppearances", 0), ("hr", "homeRuns", 0), ("rbi", "rbi", 0), ("sb", "stolenBases", 0)]


def norm(name: str) -> str:
    s = unicodedata.normalize("NFKD", name or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def official(year: str) -> dict:
    url = ("https://statsapi.mlb.com/api/v1/stats?stats=season&group=hitting&gameType=R"
           f"&season={year}&playerPool=ALL&sportId=1&limit=3000")
    data = requests.get(url, timeout=60).json()
    out: dict = {}
    for sp in data["stats"][0]["splits"]:
        out.setdefault(norm(sp["player"]["fullName"]), []).append(sp)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", nargs="+", default=["2026"])
    ap.add_argument("--min-pa", type=int, default=100)
    args = ap.parse_args()
    for year in args.year:
        idx = json.loads((ROOT / "docs/baseball/data/MLB" / f"{year}年" / "公式戦" / "batter_cards_numeric" / "index.json").read_text(encoding="utf-8"))["players"]
        off = official(year)
        rows, unmatched = [], 0
        for p in idx:
            if (p.get("pa") or 0) < args.min_pa:
                continue
            cands = off.get(norm(p["name"]), [])
            sp = None
            if len(cands) == 1:
                sp = cands[0]
            elif len(cands) > 1:   # 同姓同名は打席数が最も近い方
                sp = min(cands, key=lambda c: abs((c["stat"].get("plateAppearances") or 0) - (p.get("pa") or 0)))
            if not sp:
                unmatched += 1
                continue
            rows.append((p, sp["stat"]))
        print(f"\n===== {year}年: 照合 {len(rows)}人（打席{args.min_pa}以上）/ 公式に見つからない {unmatched}人 =====")
        for key, okey, tol in METRICS:
            diffs = []
            for p, st in rows:
                a, b = num(p.get(key)), num(st.get(okey))
                if a is None or b is None:
                    continue
                diffs.append((a - b, p, st))
            bad = [d for d in diffs if abs(d[0]) > tol + 1e-9]
            mean_abs = sum(abs(d[0]) for d in diffs) / len(diffs) if diffs else float("nan")
            mean_sig = sum(d[0] for d in diffs) / len(diffs) if diffs else float("nan")
            print(f"[{key}] ずれ(許容±{tol}): {len(bad)}/{len(diffs)}人  平均絶対差={mean_abs:.4f}  平均差(当サイト−公式)={mean_sig:+.4f}")
            for d, p, st in sorted(bad, key=lambda x: -abs(x[0]))[:6]:
                print(f"     {p['name']:<24} 当サイト {num(p.get(key))} / 公式 {st.get(okey)}  (差 {d:+.3f})")


if __name__ == "__main__":
    main()
