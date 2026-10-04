"""npb.jp の2024年投手成績ページの生HTMLを調べる読み取り専用スクリプト（表が解析できない原因の調査）。"""
from __future__ import annotations
import re

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; sports-dashboard-probe)"}


def main() -> None:
    for url in ("https://npb.jp/bis/2024/stats/idp1_c.html", "https://npb.jp/bis/2025/stats/idp1_c.html"):
        r = requests.get(url, headers=HEADERS, timeout=30)
        html = r.content.decode("utf-8", errors="replace")
        print("===", url, r.status_code, len(html), "<table:", html.count("<table"), "</table:", html.count("</table"))
        for m in list(re.finditer(r"防御率", html))[:2]:
            s = max(0, m.start() - 700)
            print("--- 防御率の周辺 ---")
            print(html[s:m.start() + 500].replace("\n", " ")[:1400])
        i = html.find("大瀬良")
        print("--- 大瀬良の周辺 ---")
        print(html[max(0, i - 300):i + 500].replace("\n", " "))


if __name__ == "__main__":
    main()
