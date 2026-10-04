"""npb.jp の2024年投手成績ページが解析できない原因の調査（読み取り専用）。parse_table の内部状態を出す。"""
from __future__ import annotations
import sys
sys.path.insert(0, "baseball/scripts")
import npbjp_history as nh


def main() -> None:
    soup = nh.get_html("https://npb.jp/bis/2024/stats/idp1_c.html")
    print("soup:", bool(soup))
    if not soup:
        return
    print("tables:", len(soup.find_all("table")), "ststats rows:", len(soup.find_all("tr", class_="ststats")))
    for t in soup.find_all("table"):
        rows = t.find_all("tr")
        print("table rows", len(rows), "防御率" in t.get_text())
        for tr in rows[:3]:
            print("  ", [nh._norm(c.get_text()) for c in tr.find_all(["th", "td"])][:30])
    r = nh.parse_table(soup, "pit")
    print("parsed:", len(r), r[:1])
    row = soup.find("tr", class_="ststats")
    if row is not None:
        tds = row.find_all(["th", "td"])
        print("first ststats row cells:", len(tds), [c.get_text(strip=True) for c in tds], "PIT_HEADS", len(nh.PIT_HEADS), nh.PIT_HEADS)


if __name__ == "__main__":
    main()
