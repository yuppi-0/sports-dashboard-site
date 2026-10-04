"""MLBの公式日程（statsapi.mlb.com）と突き合わせて、当サイトのデータに無いレギュラーシーズンの試合を洗い出す（読み取り専用）。

  python baseball/scripts/check_mlb_schedule_gaps.py --year 2026 2025

公式日程の終了した試合（Final）の gamePk が、docs/baseball/data/MLB/{年}年/公式戦/games/json/{日付}.json の
gameId に1つも無い試合を一覧にし、チームごとの試合数の差も出す。延期・中断された試合は別に表示する。
"""
from __future__ import annotations

import argparse
import collections
import glob
import gzip
import json
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[2]


def our_games(year: str) -> dict:
    out: dict = {}
    for f in sorted(glob.glob(str(ROOT / "docs/baseball/data/MLB" / f"{year}年" / "公式戦" / "games" / "json" / f"{year}-*.json*"))):
        d = json.load(gzip.open(f, "rt", encoding="utf-8")) if f.endswith(".gz") else json.load(open(f, encoding="utf-8"))
        for date, games in d.items():
            if isinstance(games, list):
                for g in games:
                    if isinstance(g, dict) and g.get("gameId"):
                        out[str(g["gameId"])] = (date, g.get("away"), g.get("home"))
    return out


def official(year: str) -> list:
    url = ("https://statsapi.mlb.com/api/v1/schedule?sportId=1&gameType=R"
           f"&startDate={year}-02-01&endDate={year}-12-31")
    res = requests.get(url, timeout=60)
    res.raise_for_status()
    games = []
    for day in res.json().get("dates", []):
        for g in day.get("games", []):
            games.append({
                "pk": str(g["gamePk"]), "date": g.get("officialDate") or day.get("date"),
                "state": (g.get("status") or {}).get("detailedState", ""),
                "away": g["teams"]["away"]["team"].get("abbreviation") or g["teams"]["away"]["team"].get("name"),
                "home": g["teams"]["home"]["team"].get("abbreviation") or g["teams"]["home"]["team"].get("name"),
                "resched": g.get("rescheduleDate") or g.get("rescheduledFrom"),
            })
    return games


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", nargs="+", default=["2026"])
    args = ap.parse_args()
    for year in args.year:
        ours = our_games(year)
        off = official(year)
        finals = [g for g in off if g["state"] in ("Final", "Game Over", "Completed Early")]
        missing = [g for g in finals if g["pk"] not in ours]
        not_final = [g for g in off if g["state"] not in ("Final", "Game Over", "Completed Early")]
        print(f"\n===== {year}年: 公式の終了試合 {len(finals)} / 当サイト {len(ours)} / 当サイトに無い {len(missing)} =====")
        for g in sorted(missing, key=lambda x: x["date"]):
            print(f"  無い: {g['date']} {g['away']}@{g['home']} gamePk={g['pk']} ({g['state']})")
        extra = set(ours) - {g["pk"] for g in off}
        if extra:
            print(f"  公式の日程に無い試合が当サイトにある: {len(extra)}件 例: {sorted(extra)[:5]}")
        if not_final:
            c = collections.Counter(g["state"] for g in not_final)
            print(f"  終了していない試合（延期・中止など）: {dict(c)}")
            for g in sorted(not_final, key=lambda x: x["date"])[:12]:
                print(f"    {g['date']} {g['away']}@{g['home']} {g['state']} gamePk={g['pk']}")


if __name__ == "__main__":
    main()
