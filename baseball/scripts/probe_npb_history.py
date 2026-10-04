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
    """2021接頭辞の試合IDを粗く走査して、2026年開幕（2021038622）より前の試合ページが残っているかを調べる。"""
    hits = []
    for gid in range(2021038600, 2021000000, -150):
        t = title(str(gid))
        if t != "取得失敗":
            hits.append((gid, t))
            print("hit", gid, t, flush=True)
    print("ヒット数:", len(hits))
    # 1軍以外（2軍・オープン戦など）が混ざる可能性があるので、ヒットの前後も確認する
    for gid, _ in hits[:3]:
        for d in (-3, -2, -1, 1, 2, 3):
            print(gid + d, title(str(gid + d)))


if __name__ == "__main__":
    main()
