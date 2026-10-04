"""Yahoo!プロ野球の球団別「選手一覧（成績）」ページの構造を調べる読み取り専用スクリプト。
公式の個人成績を全選手ぶん取れるか（当サイトのシーズン成績との突き合わせに使えるか）を確認する。"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def main() -> None:
    urls = []
    for tid in (1, 5):
        for kind in ("b", "p"):
            urls.append(f"{npb.BASE_URL}/teams/{tid}/memberlist?kind={kind}")
    urls.append(f"{npb.BASE_URL}/stats/individual?year=2026")
    for url in urls:
        soup = npb.get_soup(url)
        print("===", url)
        if not soup:
            print("取得失敗")
            continue
        print("title:", soup.title.text.strip() if soup.title else None)
        for t in soup.find_all("table")[:3]:
            rows = t.find_all("tr")
            print(f"--- table rows={len(rows)}")
            for tr in rows[:5]:
                print(" | ".join(c.get_text(strip=True) for c in tr.find_all(["th", "td"])))


if __name__ == "__main__":
    main()
