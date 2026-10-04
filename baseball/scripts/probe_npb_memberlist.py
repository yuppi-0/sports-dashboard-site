"""Yahoo!プロ野球の球団別「選手一覧（成績）」ページの構造を調べる読み取り専用スクリプト。
公式の個人成績を全選手ぶん取れるか（当サイトのシーズン成績との突き合わせに使えるか）を確認する。"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def main() -> None:
    """球団ページ・成績ページから、選手成績一覧につながるリンクを洗い出す（URLの形が分からないため）。"""
    import re
    seeds = [f"{npb.BASE_URL}/teams/1/top", f"{npb.BASE_URL}/stats/", f"{npb.BASE_URL}/stats/individual/",
             f"{npb.BASE_URL}/teams/1/memberlist", f"{npb.BASE_URL}/teams/1/stats"]
    for url in seeds:
        soup = npb.get_soup(url)
        print("===", url)
        if not soup:
            print("取得失敗")
            continue
        print("title:", soup.title.text.strip() if soup.title else None)
        seen = set()
        for a in soup.find_all("a", href=True):
            h = a["href"]
            if re.search(r"memberlist|/stats|/player/|individual|leaders|成績", h) and h not in seen:
                seen.add(h)
                if len(seen) <= 40:
                    print("  link:", h, "|", a.get_text(strip=True)[:20])
        for t in soup.find_all("table")[:2]:
            rows = t.find_all("tr")
            print(f"--- table rows={len(rows)}")
            for tr in rows[:4]:
                print(" | ".join(c.get_text(strip=True) for c in tr.find_all(["th", "td"])))


if __name__ == "__main__":
    main()
