"""日程ページが「試合の無い日」に何を表示するかを調べる読み取り専用スクリプト。

  python baseball/scripts/probe_npb_schedule.py ni 2026-09-28 2026-09-27 2026-09-29

指定した各日付について、日程ページの試合ID一覧、ページ内の「選択中の日付」らしき要素、
最初の試合ページの日付手掛かり（見出し・試合情報）を表示する。データは書き換えない。
"""
from __future__ import annotations
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def main() -> None:
    league, dates = sys.argv[1], sys.argv[2:]
    for d in dates:
        npb.set_league_dirs(league, d)
        for url in npb.SCHEDULE_URLS:
            soup = npb.get_soup(url)
            print(f"\n=== {d} {url}")
            if not soup:
                print("取得失敗")
                continue
            ids = npb._fetch_schedule_ids(url)[1]
            print("試合ID:", ids)
            print("title:", (soup.title.text.strip() if soup.title else None))
            for sel in ("h1", "h2", ".bb-calendarCtrl", ".bb-calendarCtrl__date", "[class*=selected]", "[class*=current]", ".bb-head01__title"):
                for el in soup.select(sel)[:3]:
                    print(f"[{sel}]", re.sub(r"\s+", " ", el.get_text(" ", strip=True))[:120])
            card = soup.select_one("#gm_card")
            print("gm_card見出し:", re.sub(r"\s+", " ", card.get_text(" ", strip=True))[:200] if card else None)
            if ids:
                g = npb.get_soup(f"{npb.BASE_URL}/game/{ids[0]}/")
                if g:
                    print("試合ページtitle:", g.title.text.strip() if g.title else None)
                    rd = g.select_one(".bb-gameRound")
                    print("bb-gameRound:", re.sub(r"\s+", " ", rd.get_text(" ", strip=True))[:150] if rd else None)
                    dt = [re.sub(r"\s+", " ", e.get_text(" ", strip=True))[:80] for e in g.select(".bb-gameDescription, .bb-gameDescription__left, time")][:4]
                    print("試合の日付手掛かり:", dt)


if __name__ == "__main__":
    main()
