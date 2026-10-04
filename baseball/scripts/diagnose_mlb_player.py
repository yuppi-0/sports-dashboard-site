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


def _sum_stats(stats: list, kind: str) -> dict:
    """同じ日の公式の複数試合（ダブルヘッダー）を合算する。"""
    out: dict = {}
    keys = (["inningsPitched", "strikeOuts", "baseOnBalls", "hits", "earnedRuns"] if kind == "pitching" else
            ["plateAppearances", "atBats", "hits", "baseOnBalls", "strikeOuts", "homeRuns", "rbi", "stolenBases"])
    for k in keys:
        if k == "inningsPitched":
            out[k] = sum(outs(x.get(k)) for x in stats)
        else:
            out[k] = sum(int(x.get(k) or 0) for x in stats)
    return out


def _sum_site(games: list, kind: str) -> dict:
    keys = ["ip", "k", "bb", "h", "er"] if kind == "pitching" else ["pa", "ab", "h", "bb", "k", "hr", "rbi"]
    out: dict = {}
    for k in keys:
        out[k] = sum(outs(g.get(k)) if k == "ip" else int(g.get(k) or 0) for g in games)
    return out


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
    # 当サイトの日付が米国の試合日そのままか、その翌日（日本時間表記）かは、公式の試合日と合う件数が多い方を採る
    first = next(iter(ids), None)
    official_dates: set = set()
    for pid in ids:
        gl0 = requests.get(f"{API}/people/{pid}/stats", params={"stats": "gameLog", "group": a.kind, "season": a.year,
                                                                 "gameType": "R"}, timeout=60).json()
        official_dates |= {sp["date"] for sp in ((gl0.get("stats") or [{}])[0].get("splits") or [])}
    site_dates = [g["date"] for _f, c in cards for g in site_games(c, a.kind, a.year)]
    shifts = {sh: sum((dt.date.fromisoformat(d) - dt.timedelta(days=sh)).isoformat() in official_dates for d in site_dates) for sh in (0, 1)}
    shift = max(shifts, key=shifts.get)
    print(f"日付のずらし（当サイトの日付−N日＝米国の試合日）: N={shift}（一致 {shifts}）")
    site = {}
    for _f, c in cards:
        for g in site_games(c, a.kind, a.year):
            us = (dt.date.fromisoformat(g["date"]) - dt.timedelta(days=shift)).isoformat()
            site.setdefault(us, []).append(g)
    seen_all: set = set()
    for pid, nm in ids.items():
        gl = requests.get(f"{API}/people/{pid}/stats", params={"stats": "gameLog", "group": a.kind, "season": a.year,
                                                                "gameType": "R"}, timeout=60).json()
        splits = (gl.get("stats") or [{}])[0].get("splits") or []
        print(f"\n=== {nm} id={pid} 公式 {len(splits)}試合 / 当サイト {sum(len(v) for v in site.values())}試合 ===")
        seen = seen_all
        missing = 0
        by_date: dict = {}
        for sp in splits:
            by_date.setdefault(sp["date"], []).append(sp)
        for d, sps in sorted(by_date.items()):
            sp = sps[0]
            st = _sum_stats([x["stat"] for x in sps], a.kind)
            seen.add(d)
            gs = site.get(d)
            if not gs:
                missing += 1
                if missing > 12:
                    continue
                print(f"  [当サイトに無い] {d} {sp.get('opponent', {}).get('name')} 公式: " +
                      (f"IP {st.get('inningsPitched')} K {st.get('strikeOuts')} BB {st.get('baseOnBalls')} H {st.get('hits')} ER {st.get('earnedRuns')}"
                       if a.kind == "pitching" else f"PA {st.get('plateAppearances')} AB {st.get('atBats')} H {st.get('hits')} BB {st.get('baseOnBalls')} SB {st.get('stolenBases')}"))
                continue
            g = _sum_site(gs, a.kind)
            if a.kind == "pitching":
                o = (st.get("inningsPitched"), st.get("strikeOuts"), st.get("baseOnBalls"), st.get("hits"), st.get("earnedRuns"))
                m = (g.get("ip"), g.get("k"), g.get("bb"), g.get("h"), g.get("er"))
            else:
                o = (st.get("plateAppearances"), st.get("atBats"), st.get("hits"), st.get("baseOnBalls"), st.get("strikeOuts"), st.get("homeRuns"), st.get("rbi"))
                m = (g.get("pa"), g.get("ab"), g.get("h"), g.get("bb"), g.get("k"), g.get("hr"), g.get("rbi"))
            if tuple(o) != tuple(m):
                print(f"  [値が違う] {d} 公式{o} / 当サイト{m}")
        if missing > 12:
            print(f"  …（当サイトに無い試合は計{missing}件、同姓同名の別人の試合を含む可能性）")
    extra = sorted(set(site) - seen_all)
    print(f"\n当サイトにあって公式（上の全選手のgameLog）に無い試合: {len(extra)}件 {extra[:10]}")


if __name__ == "__main__":
    main()
