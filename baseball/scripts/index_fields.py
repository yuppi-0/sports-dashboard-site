"""選手一覧（index.json）に載せる、打者の追加指標。

画面は、一覧の「打席／投球回 以上」の値に合わせて順位の母数をその場で付け直す（cards.html の liveRanksFor）。
そのためには、順位を付ける指標の値が一覧（index.json）に入っている必要がある。
打率・OPSなどは元から入っているので、ここでは Z-Swing%・Whiff%・Chase%・GB%・HR%・Hard-Hit%・平均EV・Sweet Spot%・xwOBA・
Pull%・WAR・OAA・盗塁成功率・スプリントスピード・走塁得点 を足す（値が無い選手・年度は入れない）。
"""
from __future__ import annotations

BATTER_INDEX_EXTRA = [
    "hr_pct", "chase_pct", "contact_pct", "z_swing_pct", "whiff_pct", "z_contact_pct", "gb_pct",
    "avg_ev", "hard_hit_pct", "sweet_spot_pct", "xwoba", "pull_pct",
    "sb_pct", "sprint_speed", "frp", "oaa",
]


def batter_index_extra(overall: dict | None) -> dict:
    """シーズンのoverall（カードJSONの seasons[年].overall）から、一覧に足す指標だけを取り出す。数値でないもの・無いものは含めない。"""
    out = {}
    for k in BATTER_INDEX_EXTRA:
        v = (overall or {}).get(k)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            out[k] = v
    return out
