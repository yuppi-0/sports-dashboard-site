"""npb.jp（NPB公式）の過去シーズン成績ページの構造を調べる読み取り専用スクリプト。"""
from __future__ import annotations
import re
import sys

import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; sports-dashboard-probe)"}


def get(url: str):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
    except Exception as e:  # noqa: BLE001
        print("  通信エラー:", e)
        return None
    print(f"  HTTP {r.status_code} {len(r.content)}バイト encoding={r.encoding}/{r.apparent_encoding}")
    if r.status_code != 200:
        return None
    return BeautifulSoup(r.content, "html.parser", from_encoding="utf-8")


def dump(url: str, links: bool = False, rows: int = 4) -> None:
    print("===", url)
    soup = get(url)
    if not soup:
        return
    print("  title:", soup.title.get_text(strip=True) if soup.title else None)
    if links:
        seen = []
        for a in soup.find_all("a", href=True):
            h = a["href"]
            if re.search(r"stats|idb|idp|games|scores|players", h) and h not in seen:
                seen.append(h)
        print("  links:", seen[:40])
    for t in soup.find_all("table")[:2]:
        trs = t.find_all("tr")
        print(f"  --- table rows={len(trs)}")
        for tr in trs[:rows]:
            print("   ", " | ".join(c.get_text(strip=True) for c in tr.find_all(["th", "td"])))


def main() -> None:
    dump("https://npb.jp/bis/2025/stats/", links=True)
    for name in ("idb1_c", "idp1_c", "idb1_p", "idp1_p"):
        dump(f"https://npb.jp/bis/2025/stats/{name}.html", links=False)
    dump("https://npb.jp/bis/2021/stats/idb1_c.html")
    dump("https://npb.jp/bis/2021/stats/idp1_p.html")


if __name__ == "__main__":
    main()
