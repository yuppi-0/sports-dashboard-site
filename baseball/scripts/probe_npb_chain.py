"""NPBの投球データ取得（index連鎖）が途中で止まる／重複する原因を調べる読み取り専用の調査用スクリプト。

  python baseball/scripts/probe_npb_chain.py 2021038710

試合の最初の打席ページ(index=0110100)から「次へ」ボタンをたどり、
  ・たどったindexの数と、同じindexを2回踏んでいないか
  ・連鎖が止まった最後のページのイニング・ナビゲーション（次へ/前へボタンのHTML）
  ・次のイニング先頭のindexを推定して直接開けるか
を表示する。データは書き換えない。
"""
from __future__ import annotations
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run as npb  # noqa: E402


def page(gid, idx):
    return npb.get_soup(f"{npb.BASE_URL}/game/{gid}/score?index={idx}")


def main() -> None:
    gid = sys.argv[1]
    idx, seen, last = "0110100", [], None
    while idx:
        soup = page(gid, idx)
        if not soup:
            print(f"取得失敗: index={idx}")
            break
        seen.append(idx)
        last = (idx, soup)
        nb = soup.select_one("a#btn_next")
        idx = nb["index"] if nb and "index" in nb.attrs else None
        time.sleep(0.3)
    dup = len(seen) - len(set(seen))
    print(f"試合 {gid}: ページ数 {len(seen)} / 同じindexを再訪 {dup} 回")
    print("先頭5:", seen[:5])
    print("末尾8:", seen[-8:])
    if not last:
        return
    li, ls = last
    print(f"\n最後のindex={li} イニング表示={npb.get_text(ls, 'h4.live em')}")
    for sel in ("a#btn_next", "a#btn_prev", "#btn_next", ".nextBatter"):
        for el in ls.select(sel)[:2]:
            print(f"[{sel}]", str(el)[:300].replace("\n", " "))
    print("indexを持つリンク:", [(a.get("id"), a["index"]) for a in ls.select("a[index]")][:12])
    m = re.fullmatch(r"(\d\d)(\d)(\d\d)(\d\d)", li)
    if m:
        inn, half = int(m.group(1)), int(m.group(2))
        cands = [f"{inn + 1:02d}10100", f"{inn + 1:02d}20100"] if half == 2 else [f"{inn:02d}20100", f"{inn + 1:02d}10100"]
        for c in cands:
            sp = page(gid, c)
            title = npb.get_text(sp, "h4.live em") if sp else None
            n_tbl = len([t for t in sp.select("table.bb-splitsTable")]) if sp else 0
            print(f"候補 index={c}: {'開けない' if not sp else f'イニング表示={title} 投球表={n_tbl}'}")
            time.sleep(0.5)


if __name__ == "__main__":
    main()
