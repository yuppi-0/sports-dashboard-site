"""NPB打者・投手のシーズン成績が、Yahoo!プロ野球（公式の記録）の選手ページと合っているかを調べる（読み取り専用）。

  python baseball/scripts/check_npb_official.py --year 2026

1) 当サイトの1軍試合JSONの全試合について、Yahoo!の試合ページ（/npb/game/{ID}/stats）から選手ページのリンク（選手ID）を集める
2) 各選手ページ（/npb/player/{ID}/top）の「今季成績」の行を読む
3) docs/baseball/data/プロ野球/{年}年/1軍/レギュラーシーズン/{batter,pitcher}_cards_numeric の値と突き合わせ、ずれを表示する
移籍した選手（複数球団）は選手ページがチームごとの行になるため、ずれとして出ることがある。
"""
from __future__ import annotations

import argparse
import glob
import gzip
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
BAT_KEYS = {"games": "試合", "pa": "打席", "ab": "打数", "h": "安打", "hr": "本塁打", "rbi": "打点", "avg": "打率",
            "obp": "出塁率", "slg": "長打率", "ops": "OPS"}
PIT_KEYS = {"games": "登板", "era": "防御率", "k": "奪三振", "bb": "与四球", "ip": "投球回"}


def norm(s: str) -> str:
    return re.sub(r"[\s　]", "", s or "")


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return None


def outs(ip) -> int | None:
    m = re.fullmatch(r"(\d+)(?:\.([012]))?", str(ip).strip())
    return int(m.group(1)) * 3 + int(m.group(2) or 0) if m else None


def game_ids(year: str) -> list:
    ids = []
    for f in sorted(glob.glob(str(ROOT / "docs/baseball/data/プロ野球" / f"{year}年" / "1軍" / "レギュラーシーズン" / "games" / "json" / "*.json*"))):
        d = json.load(gzip.open(f, "rt", encoding="utf-8")) if f.endswith(".gz") else json.load(open(f, encoding="utf-8"))
        for v in d.values():
            if isinstance(v, list):
                ids += [str(g["gameId"]) for g in v if isinstance(g, dict) and g.get("gameId")]
    return ids


def player_ids(gids: list) -> set:
    out: set = set()
    for i, gid in enumerate(gids):
        soup = npb.get_soup(f"{npb.BASE_URL}/game/{gid}/stats")
        if soup:
            for a in soup.find_all("a", href=True):
                m = re.search(r"/npb/player/(\d+)/", a["href"])
                if m:
                    out.add(m.group(1))
        if i % 100 == 0:
            print(f"  試合ページ {i}/{len(gids)}: 選手 {len(out)}人", flush=True)
        time.sleep(0.15)
    return out


def season_rows(soup) -> dict:
    """選手ページの表から、今季の打撃/投手の成績行を {見出し: 値} で返す（打撃: 2段の表を結合）。"""
    res = {"bat": {}, "pit": {}}
    for t in soup.find_all("table"):
        rows = [[c.get_text(strip=True) for c in tr.find_all(["th", "td"])] for tr in t.find_all("tr")]
        if len(rows) < 2:
            continue
        head = [norm(h) for h in rows[0]]
        if "打率" in head and "試合" in head and not res["bat"]:
            res["bat"].update(dict(zip(head, rows[1])))
            if len(rows) >= 4:
                res["bat"].update(dict(zip([norm(h) for h in rows[2]], rows[3])))
        elif "防御率" in head and "登板" in head and not res["pit"]:
            res["pit"].update(dict(zip(head, rows[1])))
    return res


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="2026")
    ap.add_argument("--max-games", type=int, default=0, help="試験用: 先頭からこの試合数だけで選手IDを集める")
    args = ap.parse_args()
    base = ROOT / "docs/baseball/data/プロ野球" / f"{args.year}年" / "1軍" / "レギュラーシーズン"
    cards = {"bat": {}, "pit": {}}
    for kind, d in (("bat", "batter_cards_numeric"), ("pit", "pitcher_cards_numeric")):
        for p in json.load(open(base / d / "index.json", encoding="utf-8"))["players"]:
            card = json.load(gzip.open(base / d / f"{p['id']}.json.gz", "rt", encoding="utf-8"))
            cards[kind].setdefault(norm(card["name"]), []).append(card)
    gids = game_ids(args.year)
    if args.max_games:
        gids = gids[:args.max_games]
    print(f"試合 {len(gids)} 件から選手IDを集めます")
    pids = sorted(player_ids(gids))
    print(f"選手ID {len(pids)}人 → 選手ページを確認します")
    stat = {"bat": [0, 0, 0], "pit": [0, 0, 0]}   # 照合, 不一致, カード無し
    for i, pid in enumerate(pids):
        soup = npb.get_soup(f"{npb.BASE_URL}/player/{pid}/top")
        time.sleep(0.15)
        if not soup:
            continue
        name = norm((soup.title.text if soup.title else "").split(" - ")[0])
        rows = season_rows(soup)
        for kind, keys in (("bat", BAT_KEYS), ("pit", PIT_KEYS)):
            row = rows[kind]
            if not row:
                continue
            cands = cards[kind].get(name, [])
            if not cands:
                stat[kind][2] += 1
                continue
            g_official = num(row.get("試合" if kind == "bat" else "登板"))
            card = min(cands, key=lambda c: abs((c.get("games") or 0) - (g_official or 0)))
            diffs = []
            for k, jp in keys.items():
                off = row.get(jp)
                if off is None:
                    continue
                mine = card.get(k if k != "ip" else "innings")
                if k == "ip":
                    a, b = outs(mine), outs(off)
                    if a is not None and b is not None and a != b:
                        diffs.append(f"投球回 当サイト{mine}/公式{off}")
                    continue
                o, m = num(off), num(mine)
                if o is None or m is None:
                    continue
                tol = 0.0011 if k in ("avg", "obp", "slg", "ops") else (0.011 if k == "era" else 0)
                if abs(o - m) > tol:
                    diffs.append(f"{jp} 当サイト{mine}/公式{off}")
            stat[kind][0] += 1
            if diffs:
                stat[kind][1] += 1
                print(f"  [{'打者' if kind == 'bat' else '投手'}] {name} (ID {pid}): " + ", ".join(diffs))
        if i % 100 == 0:
            print(f"  選手ページ {i}/{len(pids)}", flush=True)
    for kind, label in (("bat", "打者"), ("pit", "投手")):
        c, b, n = stat[kind]
        print(f"\n===== {label}: 照合 {c}人 / ずれ {b}人 / サイトにカードが無い {n}人 =====")


if __name__ == "__main__":
    main()
