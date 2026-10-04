"""保存済みのNPB試合の箱スコアが、Yahoo!の現在のページと違っていないかを調べる（読み取り専用）。

  python baseball/scripts/check_npb_stale_games.py --year 2026 [--days 400]

試合直後に取得した箱スコアは、後日の公式記録の訂正（安打↔失策の判定変更、打点・自責点の訂正など）が反映されない。
保存済みの全試合について Yahoo! の成績ページを取り直して比べ、違う試合・選手・項目を一覧にする。データは書き換えない。
"""
from __future__ import annotations

import argparse
import datetime
import re
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402

BASE = Path(npb.BASE_DATA_DIR) if hasattr(npb, "BASE_DATA_DIR") else Path("data/baseball/プロ野球")
BAT_COLS = ["打数", "安打", "打点", "本塁打", "四球", "死球", "三振", "犠打", "盗塁", "失策"]
PIT_COLS = ["投球回", "打者", "被安打", "被本塁打", "奪三振", "与四球", "与死球", "失点", "自責点"]


def _cells(df: pd.DataFrame) -> list:
    return [c for c in df.columns if re.fullmatch(r"\d+回", str(c))]


def _key(r) -> tuple:
    return (str(r.get("チーム")), str(r.get("選手名")))


def diff_sheet(old: pd.DataFrame, new: pd.DataFrame, cols: list, with_cells: bool) -> list:
    out = []
    if old is None or new is None or old.empty or new.empty:
        return out
    use = [c for c in cols if c in old.columns and c in new.columns]
    if with_cells:
        use += [c for c in _cells(old) if c in new.columns]
    nmap = {_key(r): r for _, r in new.iterrows()}
    for _, r in old.iterrows():
        k = _key(r)
        n = nmap.get(k)
        if n is None:
            out.append((k, "行", "保存あり", "現在のページに無い"))
            continue
        for c in use:
            a, b = str(r.get(c)).strip(), str(n.get(c)).strip()
            if a == b or (a in ("", "nan") and b in ("", "nan")):
                continue
            try:   # xlsx保存で「6」が「6.0」になる等の数値表記の違いは同じ値として扱う
                if float(a) == float(b):
                    continue
            except ValueError:
                pass
            out.append((k, c, a, b))
    return out


def collect_games(root: Path, since: str | None = None, until: str | None = None) -> list:
    """保存済みの終了試合を [(日付, 試合ID, ホーム, 打撃成績, 投手成績)] で返す。since/until（YYYY-MM-DD）で日付を絞れる。"""
    games = []
    if not root.is_dir():
        return games
    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        if (since and d.name < since) or (until and d.name > until):
            continue
        f = d / f"all_games_{d.name}.xlsx"
        if not f.exists():
            continue
        try:
            info = pd.read_excel(f, sheet_name="試合基本情報")
            bat = pd.read_excel(f, sheet_name="打撃成績")
            pit = pd.read_excel(f, sheet_name="投手成績")
        except Exception:  # noqa: BLE001
            continue
        for _, r in info.iterrows():
            if "終了" not in str(r.get("試合状態", "")):
                continue
            gid = int(r["試合ID"])
            games.append((d.name, gid, str(r.get("ホームチーム")), bat[bat["試合ID"] == gid], pit[pit["試合ID"] == gid]))
    return games


def compare_game(gid: int, home: str, bat: pd.DataFrame, pit: pd.DataFrame) -> list:
    """保存済みの箱スコアと現在のページの違いを [("打"/"投", (チーム,選手), 項目, 保存, 現在)] で返す。ページが取れなければ空。"""
    soup = npb.get_soup(f"{npb.BASE_URL}/game/{gid}/stats")
    time.sleep(0.2)
    if not soup:
        return []
    bh, br = npb._parse_batter_stats(soup, str(gid), home)
    ph, pr = npb._parse_pitcher_stats(soup, str(gid), home)
    nb = pd.DataFrame(br, columns=bh) if br else pd.DataFrame()
    np_ = pd.DataFrame(pr, columns=ph) if pr else pd.DataFrame()
    return [("打", *d) for d in diff_sheet(bat, nb, BAT_COLS, True)] + [("投", *d) for d in diff_sheet(pit, np_, PIT_COLS, False)]


def changed_dates(root: Path, since: str, until: str) -> list:
    """since〜until の保存済み試合のうち、現在のページと違う試合のある日付（後日の公式記録の訂正が未反映の日）を返す。"""
    out = set()
    for date, gid, home, bat, pit in collect_games(root, since, until):
        if compare_game(gid, home, bat, pit):
            out.add(date)
    return sorted(out)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="2026")
    ap.add_argument("--base", default=str(BASE))
    ap.add_argument("--limit", type=int, default=0, help="試験用: 先頭からこの試合数だけ調べる")
    args = ap.parse_args()
    root = Path(args.base) / f"{args.year}年" / "1軍" / "レギュラーシーズン" / "raw"
    games = collect_games(root)
    if args.limit:
        games = games[:args.limit]
    print(f"保存済みの終了試合 {len(games)} 件を現在のページと比べます")
    changed = 0
    for i, (date, gid, home, bat, pit) in enumerate(games):
        diffs = compare_game(gid, home, bat, pit)
        if diffs:
            changed += 1
            print(f"[違い] {date} 試合ID {gid} ({home}): {len(diffs)}件")
            for kind, (team, name), col, a, b in diffs[:8]:
                print(f"    {kind} {team} {name} {col}: 保存 {a} / 現在 {b}")
        if i % 100 == 0:
            print(f"  ... {i}/{len(games)}", flush=True)
    print(f"\n===== 現在のページと違う試合: {changed}/{len(games)} 件 =====")


if __name__ == "__main__":
    main()
