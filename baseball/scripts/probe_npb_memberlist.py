"""Yahoo!プロ野球の球団別「選手一覧（成績）」ページの構造を調べる読み取り専用スクリプト。
公式の個人成績を全選手ぶん取れるか（当サイトのシーズン成績との突き合わせに使えるか）を確認する。"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def main() -> None:
    """公式の個人成績ページ（リーグ別の一覧と選手ページ）の表の形を調べる。"""
    import re
    urls = [f"{npb.BASE_URL}/stats/batter?gameKindId=1&type=avg",
            f"{npb.BASE_URL}/stats/batter?gameKindId=1&type=avg&page=2",
            f"{npb.BASE_URL}/stats/pitcher?gameKindId=1&type=era",
            f"{npb.BASE_URL}/player/2000051/top",
            f"{npb.BASE_URL}/player/2000051/stats"]
    for url in urls:
        soup = npb.get_soup(url)
        print("===", url)
        if not soup:
            print("取得失敗")
            continue
        print("title:", soup.title.text.strip() if soup.title else None)
        pg = [a["href"] for a in soup.find_all("a", href=True) if re.search(r"[?&]page=|規定|all", a["href"])]
        print("  paging/links:", pg[:8])
        for t in soup.find_all("table")[:4]:
            rows = t.find_all("tr")
            print(f"--- table rows={len(rows)}")
            for tr in rows[:4]:
                print(" | ".join(c.get_text(strip=True) for c in tr.find_all(["th", "td"])))


if __name__ == "__main__":
    main()
