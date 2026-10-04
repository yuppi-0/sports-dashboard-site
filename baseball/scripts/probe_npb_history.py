"""過去シーズンのNPB試合をYahoo!から取得できるか調べる読み取り専用スクリプト。
試合IDのいくつかのページのタイトル（年月日・対戦）と、過去の日付の日程ページの表示を確認する。"""
from __future__ import annotations
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def title(gid: str) -> str:
    soup = npb.get_soup(f"{npb.BASE_URL}/game/{gid}/top")
    if not soup or not soup.title:
        return "取得失敗"
    return soup.title.text.strip()[:70]


def main() -> None:
    print("== 試合IDのサンプル（2021接頭辞）")
    for gid in [2021000001, 2021002000, 2021005000, 2021010000, 2021015000, 2021020000, 2021025000, 2021030000,
                2021034000, 2021036000, 2021038000, 2021038622]:
        print(gid, title(str(gid)))
    print("== 他の接頭辞")
    for pre in (2018, 2019, 2020, 2022, 2023, 2024, 2025):
        gid = f"{pre}000001"
        gid = f"{pre}0000001"[:10]
        print(gid, title(gid))
    print("== 過去日付の日程ページ")
    for d in ("2025-06-01", "2024-06-01", "2023-06-01", "2021-06-01"):
        soup = npb.get_soup(f"{npb.BASE_URL}/schedule/first/all?date={d}")
        if not soup:
            print(d, "取得失敗")
            continue
        h = soup.select_one(".bb-head01__title")
        ids = [m.group(1) for a in soup.select("#gm_card a[href]") for m in [re.search(r"/npb/game/(\d+)/", a["href"])] if m]
        print(d, "見出し:", h.get_text(strip=True) if h else None, "試合ID:", ids[:4])


if __name__ == "__main__":
    main()
