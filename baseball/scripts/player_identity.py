"""同じシーズンに同姓同名の別人がいる選手（例: Max Muncy=LAD/ATH、Luis García、Will Smith）を別カードに分けるヘルパー。

試合JSONの選手エントリは名前しか持たないため、名前だけで集計すると別人の成績が1枚のカードに合算されてしまう
（公式との突き合わせで、打席・打点・投球回が数十〜200単位でずれていた）。
同じ日に別々のチームで同じ名前の選手が出場している場合は別人とみなし、名前を「Max Muncy (ATH)」の形に変える。
トレードで移籍した選手は同じ日に2チームで出場しないので、これまでどおり1枚のカードにまとまる。
"""
from __future__ import annotations
import unicodedata


def fold_name(name: str) -> str:
    s = unicodedata.normalize("NFKD", name or "")
    return "".join(ch for ch in s if not unicodedata.combining(ch)).lower().strip()


def _normalize(name: str) -> str:
    name = (name or "").strip()
    if "," in name:
        last, first = (x.strip() for x in name.split(",", 1))
        if last and first:
            return f"{first} {last}"
    return name


def disambiguate_same_name(all_data: dict, key: str) -> dict:
    """all_data[日付] = [試合dict, ...]（試合dict[key]["home"/"away"] = 選手エントリのリスト）を調べ、
    同じ日に別チームで同名の選手がいる名前だけ、エントリのnameを「名前 (チーム)」に書き換える（in-place）。
    key は "pitchers" か "batters"。書き換えた {名前の小文字キー: [チーム,...]} を返す。"""
    teams_by_date: dict = {}   # (名前キー, 日付) -> {チーム}
    for date, games in all_data.items():
        if date == "highlights" or str(date).startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict) or not isinstance(g.get(key), dict):
                continue
            for side in ("home", "away"):
                team = g.get(side)
                for p in (g[key].get(side) or []):
                    if isinstance(p, dict) and p.get("name") and team:
                        teams_by_date.setdefault((fold_name(_normalize(p["name"])), date), set()).add(team)
    conflict: dict = {}
    for (nk, _date), teams in teams_by_date.items():
        if len(teams) >= 2:
            conflict.setdefault(nk, set()).update(teams)
    if not conflict:
        return {}
    for date, games in all_data.items():
        if date == "highlights" or str(date).startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict) or not isinstance(g.get(key), dict):
                continue
            for side in ("home", "away"):
                team = g.get(side)
                for p in (g[key].get(side) or []):
                    if not isinstance(p, dict) or not p.get("name") or not team:
                        continue
                    nk = fold_name(_normalize(p["name"]))
                    if nk in conflict:
                        p["name"] = f"{_normalize(p['name'])} ({team})"
    return {k: sorted(v) for k, v in conflict.items()}
