"""MLB選手の試合ごとの成績を、MLB公式（statsapi）のgameLogと当サイトのカードのgame_logで突き合わせて、
どの試合がずれているかを出す（読み取り専用）。シーズン合計がずれた選手の原因調査用。

  python baseball/scripts/diagnose_mlb_player.py --kind pitching --year 2021 "Joe Smith"
  python baseball/scripts/diagnose_mlb_player.py --kind hitting --year 2024 "Danny Jansen"

当サイトの日付は米国の試合日の翌日（日本時間表記）なので、1日引いて公式の日付と突き合わせる。
"""
from __future__ import annotations

import argparse
import datetime as dt
import gzip
import json
import re
import unicodedata
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[2]
API = "https://statsapi.mlb.com/api/v1"


def norm(name: str) -> str:
    s = unicodedata.normalize("NFKD", name or "")
    return re.sub(r"[^a-z0-9]", "", "".join(c for c in s if not unicodedata.combining(c)).lower())


def outs(ip) -> int:
    a, _, b = str(ip or "0").partition(".")
    return int(a) * 3 + int(b or 0)


def load_card(year: str, kind: str, name: str):
    sub = "pitcher" if kind == "pitching" else "batter"
    d = ROOT / "docs/baseball/data/MLB" / f"{year}年" / "公式戦" / f"{sub}_cards_numeric"
    for f in d.glob("*.json.gz"):
        if f.name == "index.json":
            continue
        try:
            c = json.load(gzip.open(f))
        except Exception:  # noqa: BLE001
            continue
        if norm(c.get("name", "")) == norm(name):
            yield f.name, c


def site_games(card: dict, kind: str, year: str) -> list:
    gl = card.get("game_log") if kind == "pitching" else (card.get("seasons", {}).get(year, {}).get("game_log"))
    return gl or []


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("name")
    ap.add_argument("--kind", choices=["pitching", "hitting"], required=True)
    ap.add_argument("--year", required=True)
    a = ap.parse_args()
    r = requests.get(f"{API}/stats", params={"stats": "season", "group": a.kind, "gameType": "R", "season": a.year,
                                             "playerPool": "ALL", "sportId": 1, "limit": 3000}, timeout=60).json()
    ids = {}
    for sp in r["stats"][0]["splits"]:
        p = sp["player"]
        if norm(p["fullName"]) == norm(a.name):
            ids[p["id"]] = p["fullName"]
    print("公式の該当選手:", ids)
    cards = list(load_card(a.year, a.kind, a.name))
    print("当サイトのカード:", [(f, c.get("team")) for f, c in cards])
    site = {}
    for _f, c in cards:
        for g in site_games(c, a.kind, a.year):
            us = (dt.date.fromisoformat(g["date"]) - dt.timedelta(days=1)).isoformat()
            site.setdefault(us, []).append(g)
    for pid, nm in ids.items():
        gl = requests.get(f"{API}/people/{pid}/stats", params={"stats": "gameLog", "group": a.kind, "season": a.year,
                                                                "gameType": "R"}, timeout=60).json()
        splits = (gl.get("stats") or [{}])[0].get("splits") or []
        print(f"\n=== {nm} id={pid} 公式 {len(splits)}試合 / 当サイト {sum(len(v) for v in site.values())}試合 ===")
        seen = set()
        for sp in splits:
            d, st = sp["date"], sp["stat"]
            seen.add(d)
            gs = site.get(d)
            if not gs:
                print(f"  [当サイトに無い] {d} {sp.get('opponent', {}).get('name')} 公式: " +
                      (f"IP {st.get('inningsPitched')} K {st.get('strikeOuts')} BB {st.get('baseOnBalls')} H {st.get('hits')} ER {st.get('earnedRuns')}"
                       if a.kind == "pitching" else f"PA {st.get('plateAppearances')} AB {st.get('atBats')} H {st.get('hits')} BB {st.get('baseOnBalls')} SB {st.get('stolenBases')}"))
                continue
            g = gs[0]
            if a.kind == "pitching":
                o = (outs(st.get("inningsPitched")), st.get("strikeOuts"), st.get("baseOnBalls"), st.get("hits"), st.get("earnedRuns"))
                m = (outs(g.get("ip")), g.get("k"), g.get("bb"), g.get("h"), g.get("er"))
            else:
                o = (st.get("plateAppearances"), st.get("atBats"), st.get("hits"), st.get("baseOnBalls"), st.get("strikeOuts"), st.get("homeRuns"), st.get("rbi"))
                m = (g.get("pa"), g.get("ab"), g.get("h"), g.get("bb"), g.get("k"), g.get("hr"), g.get("rbi"))
            if tuple(o) != tuple(m):
                print(f"  [値が違う] {d} 公式{o} / 当サイト{m}")
        for d in sorted(set(site) - seen):
            print(f"  [公式に無い] {d}")


if __name__ == "__main__":
    main()
