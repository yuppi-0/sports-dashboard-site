"""選手の同一性（同じ人は1枚、別人は別カード）をそろえるヘルパー。

[選手IDがあるとき] 試合JSONの選手エントリに MLBAM ID（"id"）があれば、IDで同一選手を判定する。
  ・表記が途中で変わる選手（Statcastの名前が "Zach" → "Zac" に変わる、アクセント有無の違い）は、最新の表記に統一して1枚にまとめる
  ・同じ表記の別人（Max Muncy=LAD/ATH など）は、IDが違うので「名前 (チーム)」に分ける
[選手IDが無いとき（古いデータ・NPB）] 下の日付の重なりで判定する。

同じシーズンに同姓同名の別人がいる選手（例: Max Muncy=LAD/ATH、Luis García、Will Smith）を別カードに分けるヘルパー。

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


def same_name_conflicts(all_data: dict, key: str) -> dict:
    """同じ日に別チームで同名の選手がいる名前 {名前の小文字キー: {チーム,...}}（書き換えはしない）。"""
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
    return conflict


def disambiguate_same_name(all_data: dict, key: str, extra_conflicts: dict | None = None) -> dict:
    """all_data[日付] = [試合dict, ...]（試合dict[key]["home"/"away"] = 選手エントリのリスト）を調べ、
    同じ日に別チームで同名の選手がいる名前だけ、エントリのnameを「名前 (チーム)」に書き換える（in-place）。
    key は "pitchers" か "batters"。書き換えた {名前の小文字キー: [チーム,...]} を返す。
    extra_conflicts: 反対側（投手↔打者）で同姓同名の別人と分かった {名前キー: {チーム}}。投手側だけでは同じ日に
    重ならない別人（例: 投手のオスナ=ソフトバンクと、野手登板が1試合だけのヤクルトのオスナ）も、これで分けられる。"""
    conflict = same_name_conflicts(all_data, key)
    for nk, teams in (extra_conflicts or {}).items():
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


def _entries(all_data: dict, key: str):
    """(日付, 試合dict, チーム, 選手エントリ) を順に返す。"""
    for date, games in all_data.items():
        if date == "highlights" or str(date).startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict) or not isinstance(g.get(key), dict):
                continue
            for side in ("home", "away"):
                for p in (g[key].get(side) or []):
                    if isinstance(p, dict) and p.get("name"):
                        yield date, g, g.get(side), p


def canonicalize_by_id(all_data: dict, key: str) -> dict:
    """選手IDを持つエントリの名前を、IDごとに最新の表記へ統一し、同じ表記の別ID（別人）は「名前 (チーム)」に分ける（in-place）。
    IDを持たないエントリには触らない。分けた名前の {名前キー: [チーム,...]} を返す。"""
    latest: dict = {}     # id -> (日付, 名前)
    for date, _g, _team, p in _entries(all_data, key):
        pid = p.get("id")
        if pid is None:
            continue
        nm = _normalize(p["name"])
        cur = latest.get(pid)
        if cur is None or (date, sum(ord(c) > 127 for c in nm)) >= (cur[0], sum(ord(c) > 127 for c in cur[1])):
            latest[pid] = (date, nm)
    if not latest:
        return {}
    ids_by_fold: dict = {}
    for pid, (_d, nm) in latest.items():
        ids_by_fold.setdefault(fold_name(nm), set()).add(pid)
    clash = {k for k, v in ids_by_fold.items() if len(v) >= 2}
    teams_of: dict = {}
    for _date, _g, team, p in _entries(all_data, key):
        pid = p.get("id")
        if pid is None:
            continue
        nm = latest[pid][1]
        if fold_name(nm) in clash and team:
            p["name"] = f"{nm} ({team})"
            teams_of.setdefault(fold_name(nm), set()).add(team)
        else:
            p["name"] = nm
    return {k: sorted(v) for k, v in teams_of.items()}
