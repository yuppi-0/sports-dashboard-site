#!/usr/bin/env python3
"""
run.py（whoscored_player_stats.py）が出力した「結合」xlsx から、サッカー選手カード用の数値JSONを作る。

入力（データリポジトリ sports-dashboard-data の soccer/）:
    <開幕年>年/all_leagues_players_<開幕年>.xlsx … 選手のシーズン集計（全リーグ。列 league / season / venue = all・home・away、_順位・_順位_母数つき）
    <開幕年>年/all_leagues_matches_<開幕年>.xlsx … 選手×試合（同上）
      ※ run.py の output/結合/ に出る「2025年」「2026年」フォルダをそのまま置く（1ファイル1シート。matches は約19MBで、GitHubのブラウザ上げの25MiB以内）。
      ※ フォルダ名は問わない。ファイル名の年（_2025 など）で判断する。

出力（サイトリポジトリ sports-dashboard-site）:
    docs/soccer/data/manifest.json                                   … 取得できるリーグ・シーズンの一覧
    docs/soccer/data/<League>/<開幕年>/cards_numeric/index.json      … 選手一覧（一覧・並び替え用の軽い情報）
    docs/soccer/data/<League>/<開幕年>/cards_numeric/<id>.json       … 選手ごとのカード数値
    docs/soccer/data/<League>/<開幕年>/cards_llm/batches/index.json  … LLM解釈JSON（手動で置く）の一覧。このスクリプトが作り直す
      ※ LLM解釈JSON は「{所属クラブ}_{FW|MF|DF|GK}.json」（プロンプト4.2）を batches/ に置く。

使い方:
    python soccer/scripts/build_soccer_cards.py --data-dir ../sports-dashboard-data/soccer --out docs/soccer/data
    python soccer/scripts/build_soccer_cards.py ... --leagues Premier LaLiga --seasons 2026 --min-minutes 180

依存: pandas numpy openpyxl
"""
import argparse
import hashlib
import json
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd

LEAGUE_LABELS = {
    "Premier": "プレミアリーグ", "LaLiga": "ラ・リーガ", "Bundesliga": "ブンデスリーガ",
    "SerieA": "セリエA", "Ligue1": "リーグ・アン",
}
VENUES = ("all", "home", "away")

# ---- カードに載せる指標の定義 -------------------------------------------------------
# (列名, 表示名, 単位, 良い方向)   ※ 良い方向: high=大きいほど良い / low=小さいほど良い
# 列が入力に無い指標（例: Understatを結合していないときのxG・xA）は自動で省く。
# 付録A（コア指標）・付録B（プレータイプ）は RANK_SPECS（run.py）と揃えてある。
CORE = {
    "FW": [("goals_p90", "ゴール/90", "", "high"), ("xg_diff_p90", "xG差分/90", "", "high"),
           ("shot_on_target_pct", "シュート枠内率", "%", "high"), ("touches_att_pen_pct", "PA内タッチ率", "%", "high"),
           ("take_on_pct", "ドリブル成功率", "%", "high"), ("def_actions_att_third_p90", "敵陣守備アクション/90", "", "high")],
    "MF": [("pass_pct", "パス成功率", "%", "high"), ("progressive_passes_p90", "プログレッシブパス/90", "", "high"),
           ("us_xa_p90", "xA/90", "", "high"), ("key_passes_p90", "キーパス/90", "", "high"),
           ("tackle_win_pct", "タックル成功率", "%", "high"), ("interceptions_p90", "インターセプト/90", "", "high")],
    "CB": [("tackle_win_pct", "タックル成功率", "%", "high"), ("interceptions_p90", "インターセプト/90", "", "high"),
           ("clearances_p90", "クリア/90", "", "high"), ("aerial_win_pct", "空中戦勝率", "%", "high"),
           ("dribbled_past_pct", "被ドリブル率", "%", "low"), ("pass_pct", "パス成功率", "%", "high")],
    "SB": [("tackle_win_pct", "タックル成功率", "%", "high"), ("dribbled_past_pct", "被ドリブル率", "%", "low"),
           ("crosses_p90", "クロス/90", "", "high"), ("progressive_carries_p90", "プログレッシブラン/90", "", "high"),
           ("key_passes_p90", "キーパス/90", "", "high"), ("duel_win_pct", "デュエル勝率", "%", "high")],
    "GK": [("gk_save_pct", "セーブ率", "%", "high"), ("goals_conceded_p90", "失点/90", "", "low"),
           ("clean_sheet_pct", "クリーンシート率", "%", "high"), ("pass_pct", "パス成功率", "%", "high"),
           ("long_ball_pct", "ロングボール成功率", "%", "high"), ("gk_sweeper_actions_p90", "エリア外対応/90", "", "high")],
}

# プレータイプ。name はプロンプト付録Bの名前と完全一致させる（LLM解釈JSONの play_evaluations[].name と突き合わせる）。
# count=そのプレータイプの回数の列、share_of=使用割合の分母（列名 / 列名のリスト=合計）
_ZONES = [("Def 3rd", "自陣3分の1", "def_third"), ("Mid 3rd", "中央3分の1", "mid_third"), ("Att 3rd", "敵陣3分の1", "att_third")]
_DF_TYPES = {
    "count_label": "守備アクション", "share_of": [f"def_actions_{z}" for _, _, z in _ZONES],
    "types": [{"name": n, "label": l, "count": f"def_actions_{z}", "metrics": [
        (f"tackles_{z}_p90", "タックル/90", "", "high"), (f"tackle_win_pct_{z}", "タックル成功率", "%", "high"),
        (f"interceptions_{z}_p90", "インターセプト/90", "", "high")]} for n, l, z in _ZONES],
}
PLAYTYPES = {
    "FW": {"count_label": "シュートにつながったプレー（近似）", "share_of": "sca", "types": [
        {"name": "Live-ball Pass", "label": "流れの中のパスから", "count": "sca_pass_live", "metrics": [("sca_pass_live_p90", "回数/90", "", "high")]},
        {"name": "Dead-ball Pass", "label": "セットプレーから", "count": "sca_pass_dead", "metrics": [("sca_pass_dead_p90", "回数/90", "", "high")]},
        {"name": "Take-On", "label": "ドリブル突破から", "count": "sca_take_on", "metrics": [("sca_take_on_p90", "回数/90", "", "high")]},
        {"name": "Shot", "label": "こぼれ球・セカンドから", "count": "sca_shot", "metrics": [("sca_shot_p90", "回数/90", "", "high")]},
        {"name": "Fouls Drawn", "label": "ファウル奪取から", "count": "sca_foul_won", "metrics": [("sca_foul_won_p90", "回数/90", "", "high")]},
        {"name": "Defensive Action", "label": "守備からの転化", "count": "sca_defensive", "metrics": [("sca_defensive_p90", "回数/90", "", "high")]},
    ]},
    "MF": {"count_label": "パス", "share_of": "passes", "types": [
        {"name": "Short", "label": "短距離", "count": "passes_short", "metrics": [("pass_short_pct", "成功率", "%", "high")]},
        {"name": "Medium", "label": "中距離", "count": "passes_medium", "metrics": [("pass_medium_pct", "成功率", "%", "high")]},
        {"name": "Long", "label": "長距離", "count": "passes_long", "metrics": [("pass_long_pct", "成功率", "%", "high")]},
        {"name": "Dead", "label": "セットプレー", "count": "passes_dead", "metrics": [("passes_dead_pct", "成功率", "%", "high")]},
        {"name": "FK", "label": "フリーキック", "count": "passes_fk", "metrics": [("passes_fk_pct", "成功率", "%", "high")]},
        {"name": "TB", "label": "スルーパス", "count": "through_balls", "metrics": [("through_ball_pct", "成功率", "%", "high")]},
        {"name": "Sw", "label": "サイドチェンジ", "count": "switches", "metrics": [("switch_pct", "成功率", "%", "high")]},
        {"name": "Crs", "label": "クロス", "count": "crosses", "metrics": [("cross_pct", "成功率", "%", "high")]},
    ]},
    "CB": _DF_TYPES,
    "SB": _DF_TYPES,
}

# 試合ログに出す列（共通列のあとに、集団ごとの列）
GAMELOG_COMMON = [("minutes", "出場"), ("rating", "評価")]
GAMELOG_EXTRA = {
    "FW": [("goals", "得点"), ("shots", "シュート"), ("shots_on_target", "枠内"), ("key_passes", "キーパス")],
    "MF": [("goals", "得点"), ("passes", "パス"), ("pass_pct", "パス%"), ("key_passes", "キーパス"), ("tackles", "タックル"), ("interceptions", "INT")],
    "CB": [("tackles", "タックル"), ("interceptions", "INT"), ("clearances", "クリア"), ("aerials_won", "空中戦勝")],
    "SB": [("tackles", "タックル"), ("interceptions", "INT"), ("crosses", "クロス"), ("key_passes", "キーパス")],
    "GK": [("saves", "セーブ"), ("goals_conceded", "失点"), ("clean_sheets", "無失点"), ("pass_pct", "パス%")],
}

# 一覧の並び替えに使う値（venue=all の行から取る）
INDEX_SORT_COLS = ["rating_avg", "minutes", "matches", "goals_p90", "key_passes_p90",
                   "progressive_passes_p90", "tackles_p90", "interceptions_p90"]


# ---- 小さな道具 ----------------------------------------------------------------------
def num(v):
    """NaN/欠損は None、整数値の float は int、それ以外は float にする（JSONに載せる用）。"""
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA:
        return None
    if isinstance(v, (np.integer, int)):
        return int(v)
    if isinstance(v, (np.floating, float)):
        f = float(v)
        return int(f) if f.is_integer() and abs(f) < 1e9 else round(f, 3)
    return v


def text(v):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA else str(v)


def slug(s):
    """クラブ名などをファイル名に使える英数字にする（別の名前が同じにならないよう、短いハッシュを付ける）。"""
    a = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    a = re.sub(r"[^0-9A-Za-z]+", "-", a).strip("-").lower()
    return f"{a}-{hashlib.md5(str(s).encode()).hexdigest()[:4]}" if a else hashlib.md5(str(s).encode()).hexdigest()[:8]


def season_label(year):
    return f"{year}-{str(year + 1)[2:]}"


def write_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


# ---- カードの部品 --------------------------------------------------------------------
def venue_rows(g):
    """選手×クラブの、all/home/away の3行を {venue: Series} にする。"""
    return {v: r for v, r in ((row["venue"], row) for _, row in g.iterrows()) if v in VENUES}


def cell(rows, venue, col):
    r = rows.get(venue)
    if r is None or col not in r.index:
        return None
    v = num(r[col])
    if v is None:
        return None
    rank, pop = num(r.get(col + "_順位")), num(r.get(col + "_順位_母数"))
    return {"v": v, "rank": rank, "pop": pop}


def metric_entry(rows, spec, columns):
    col, label, unit, better = spec
    if col not in columns:
        return None
    entry = {"key": col, "label": label, "unit": unit, "better": better}
    for v in VENUES:
        entry[v] = cell(rows, v, col)
    return entry if any(entry[v] for v in VENUES) else None


def playtype_entries(group, rows, columns):
    cfg = PLAYTYPES.get(group)
    if not cfg:
        return None
    share = cfg["share_of"]
    out = {"count_label": cfg["count_label"], "types": []}
    for t in cfg["types"]:
        if t["count"] not in columns:
            continue
        item = {"name": t["name"], "label": t["label"]}
        for v in VENUES:
            r = rows.get(v)
            if r is None:
                item[v] = None
                continue
            n = num(r.get(t["count"]))
            denom_cols = share if isinstance(share, list) else [share]
            d = sum((num(r.get(c)) or 0) for c in denom_cols)
            item[v] = {"count": n, "pct": round(n / d * 100, 1) if (n is not None and d) else None}
        item["metrics"] = [m for m in (metric_entry(rows, s, columns) for s in t["metrics"]) if m]
        out["types"].append(item)
    return out if out["types"] else None


def game_log(matches_g, group):
    cols_def = GAMELOG_COMMON + GAMELOG_EXTRA[group]
    cols = [(c, l) for c, l in cols_def if c in matches_g.columns]
    rows = []
    for _, r in matches_g.sort_values("date").iterrows():
        rows.append({
            "date": text(r["date"])[:10], "opponent": text(r["opponent"]), "venue": "H" if r["venue"] == "home" else "A",
            "gf": num(r["goals_for"]), "ga": num(r["goals_against"]),
            "starter": bool(num(r.get("is_starter"))), "pos": text(r.get("position_played")),
            **{c: num(r[c]) for c, _ in cols},
        })
    return [{"key": c, "label": l} for c, l in cols], rows


def build_cards(players, matches, league, year, min_minutes):
    columns = set(players.columns)
    players = players[players["position"].notna()].copy()
    players["group"] = np.where(players["position"] == "DF", players["position_detail"], players["position"])
    mkey = {k: g for k, g in matches.groupby(["team", "player_id"])}
    cards, index = [], []
    for (team, pid), g in players.groupby(["team", "player_id"]):
        rows = venue_rows(g)
        a = rows.get("all")
        if a is None or (num(a["minutes"]) or 0) < min_minutes:
            continue
        group = a["group"]
        if group not in CORE:
            continue
        cid = f"{int(pid)}_{slug(team)}"
        log_cols, log = game_log(mkey[(team, pid)], group) if (team, pid) in mkey else ([], [])
        card = {
            "id": cid, "name": a["player"], "team": team, "league": league, "season": str(year),
            "position": a["position"], "group": group,
            "matches": num(a["matches"]), "starts": num(a["starts"]), "minutes": num(a["minutes"]),
            "rating_avg": num(a.get("rating_avg")), "motm": num(a.get("motm")),
            "goals": num(a.get("goals")), "clean_sheets": num(a.get("clean_sheets")),
            "core": [m for m in (metric_entry(rows, s, columns) for s in CORE[group]) if m],
            "playtypes": playtype_entries(group, rows, columns),
            "game_log_cols": log_cols, "game_log": log,
        }
        cards.append(card)
        idx = {"id": cid, "name": a["player"], "team": team, "position": a["position"], "group": group}
        for c in INDEX_SORT_COLS:
            idx[c] = num(a.get(c)) if c in columns else None
        index.append(idx)
    return cards, index


# ---- 出力 --------------------------------------------------------------------------
def write_league_season(out_dir, league, year, cards, index):
    base = out_dir / league / str(year)
    numeric = base / "cards_numeric"
    numeric.mkdir(parents=True, exist_ok=True)
    for old in numeric.glob("*.json"):          # 閾値の変更や移籍で消えた選手のファイルを残さない
        old.unlink()
    for c in cards:
        write_json(numeric / f"{c['id']}.json", c)
    index.sort(key=lambda p: (p["team"], p["name"]))
    write_json(numeric / "index.json", {
        "league": league, "season": str(year), "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "teams": sorted({p["team"] for p in index}), "players": index})
    batches = base / "cards_llm" / "batches"    # LLM解釈JSONは手動で置く。一覧だけここで作り直す
    batches.mkdir(parents=True, exist_ok=True)
    files = sorted(f.name for f in batches.glob("*.json") if f.name != "index.json")
    write_json(batches / "index.json", {"batches": files})
    return len(cards), len(files)


def write_manifest(out_dir):
    """出力フォルダにある cards_numeric/index.json を走査して作るので、一部のリーグだけ更新しても他が消えない。"""
    leagues = {}
    for idx in sorted(out_dir.glob("*/*/cards_numeric/index.json")):
        league, year = idx.parts[-4], idx.parts[-3]
        bidx = idx.parents[1] / "cards_llm" / "batches" / "index.json"
        n_llm = len(json.loads(bidx.read_text(encoding="utf-8")).get("batches", [])) if bidx.exists() else 0
        leagues.setdefault(league, {"label": LEAGUE_LABELS.get(league, league), "seasons": []})["seasons"].append(
            {"year": year, "label": season_label(int(year)), "llm_batches": n_llm})
    for v in leagues.values():
        v["seasons"].sort(key=lambda s: s["year"], reverse=True)
    order = [k for k in LEAGUE_LABELS if k in leagues] + [k for k in leagues if k not in LEAGUE_LABELS]
    write_json(out_dir / "manifest.json", {"leagues": {k: leagues[k] for k in order}})


YEAR_FILE_RE = re.compile(r"all_leagues_(players|matches)_(\d{4})\.xlsx")


def find_sources(data_dir, seasons):
    """
    {開幕年: {"players": ファイル, "matches": ファイル}}。
    soccer/2025年/all_leagues_players_2025.xlsx など。フォルダ名は問わず、ファイル名の年で判断する（1ファイル1シート）。
    """
    cand = {}
    for f in sorted(data_dir.rglob("all_leagues_*_*.xlsx")):
        m = YEAR_FILE_RE.fullmatch(f.name)
        if m and (not seasons or int(m.group(2)) in seasons):
            cand.setdefault((int(m.group(2)), m.group(1)), []).append(f)
    src = {}
    for (year, kind), files in cand.items():
        if len(files) > 1:
            files.sort(key=lambda f: f.stat().st_mtime)
            print(f"注意: {year} の {kind} が複数あります。新しい方を使います: {files[-1]}（他: {[str(f) for f in files[:-1]]}）")
        src.setdefault(year, {})[kind] = files[-1]
    return src


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data-dir", required=True, type=Path,
                    help="年ごとのフォルダ（2025年/ など）が入っているフォルダ（soccer）。中を再帰的に探す")
    ap.add_argument("--out", required=True, type=Path, help="出力先（docs/soccer/data）")
    ap.add_argument("--leagues", nargs="*", help="対象リーグ（省略で全部）")
    ap.add_argument("--seasons", nargs="*", type=int, help="対象の開幕年（省略で全部）")
    ap.add_argument("--min-minutes", type=float, default=900, help="カードを作る出場時間の下限（既定900分）")
    args = ap.parse_args()

    if not args.data_dir.is_dir():
        sys.exit(f"フォルダが見つかりません: {args.data_dir}")
    src = find_sources(args.data_dir, args.seasons)
    if not src:
        sys.exit(f"対象のxlsxがありません: {args.data_dir}（all_leagues_players_<年>.xlsx / all_leagues_matches_<年>.xlsx）")

    for year in sorted(src, reverse=True):
        s = src[year]
        if set(s) != {"players", "matches"}:
            print(f"注意: {year} は players / matches の片方しか無いため、スキップします（あるのは {sorted(s)}）")
            continue
        players = pd.read_excel(s["players"])
        matches = pd.read_excel(s["matches"])
        for league in [l for l in LEAGUE_LABELS if l in set(players["league"])] + \
                      sorted(set(players["league"]) - set(LEAGUE_LABELS)):
            if args.leagues and league not in args.leagues:
                continue
            cards, index = build_cards(players[players["league"] == league], matches[matches["league"] == league],
                                       league, year, args.min_minutes)
            if not cards:
                print(f"[{league} {year}] 出場{args.min_minutes:.0f}分以上の選手がいないため、スキップ")
                continue
            n, nb = write_league_season(args.out, league, year, cards, index)
            print(f"[{league} {year}] カード{n}枚（LLM解釈バッチ{nb}ファイル）→ {args.out / league / str(year)}")
    write_manifest(args.out)


if __name__ == "__main__":
    main()
