"""順位表ページの構造を調べる読み取り専用スクリプト（チーム別の試合数を取得できるか確認する）。"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def main() -> None:
    for url in (f"{npb.BASE_URL}/standings/", f"{npb.BASE_URL}/standings/?year=2026"):
        soup = npb.get_soup(url)
        print("===", url)
        if not soup:
            print("取得失敗")
            continue
        for t in soup.find_all("table"):
            rows = t.find_all("tr")
            print(f"--- table rows={len(rows)} class={t.get('class')}")
            for tr in rows[:9]:
                print(" | ".join(c.get_text(strip=True) for c in tr.find_all(["th", "td"])))


if __name__ == "__main__":
    main()
