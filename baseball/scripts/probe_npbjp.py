"""npb.jp のファーム（2軍）成績ページのURLと表の形を調べる読み取り専用スクリプト。"""
from __future__ import annotations
import sys

sys.path.insert(0, "baseball/scripts")
import npbjp_history as nh


def main() -> None:
    soup = nh.get_html("https://npb.jp/bis/2024/stats/idp1_c.html")
    for a in soup.find_all("a", href=True):
        t = a.get_text(strip=True)
        if "個人" in t or "ファーム" in t or "チーム" in t:
            print("link:", t, a["href"])
    for url in ("https://npb.jp/bis/2024/stats/idb2_c.html", "https://npb.jp/bis/2024/stats/idp2_c.html",
                "https://npb.jp/bis/2021/stats/idp2_c.html", "https://npb.jp/bis/2025/stats/idp2_c.html"):
        s = nh.get_html(url)
        print("==", url, bool(s))
        if not s:
            continue
        print(" ststats rows:", len(s.find_all("tr", class_="ststats")), "tables:", len(s.find_all("table")))
        for t in s.find_all("table"):
            rows = t.find_all("tr")
            for tr in rows[:3]:
                print("  ", [nh._norm(c.get_text()) for c in tr.find_all(["th", "td"])][:30])
        print(" parsed:", len(nh.parse_table(s, "pit" if "idp" in url else "bat")))


if __name__ == "__main__":
    main()
