#!/usr/bin/env python3
"""
run.py（whoscored_player_stats.py v10）が出力した「結合」xlsx から、サッカー選手カード用の数値JSONを作る。

入力（データリポジトリ sports-dashboard-data の soccer/）:
    <開幕年>年/all_leagues_players_<開幕年>.xlsx          … 選手のシーズン集計（全リーグ。列 league / season / venue = all・home・away、_順位・_順位_母数つき）
    <開幕年>年/all_leagues_matches_<開幕年>.xlsx          … 選手×試合（全リーグ。1ファイルのとき）
    <開幕年>年/all_leagues_matches_<開幕年>_<League>.xlsx … 選手×試合（リーグ別。run.py --split-by-league の出力。1ファイルが25MiBを超えるとき用）
      ※ run.py の output/結合/ に出る「2025年」「2026年」フォルダをそのまま置く（1ファイル1シート）。
      ※ フォルダ名は問わない。ファイル名の年（_2025 など）で判断する。
      ※ リーグ別のファイルがあるリーグは、そちらを使う（全リーグ1ファイルが残っていても、そのリーグの分は無視する）。

出力（サイトリポジトリ sports-dashboard-site）:
    docs/soccer/data/manifest.json                                   … 取得できるリーグ・シーズンの一覧（games = 消化した試合数の目安）
    docs/soccer/data/games/<開幕年>/<League>/index.json              … 選手一覧（一覧・並び替え用の軽い情報）
    docs/soccer/data/games/<開幕年>/<League>/<id>.json               … 選手ごとのカード数値
    docs/soccer/data/llm/<開幕年>/index.json                        … LLM解釈JSON（手動で置く）の一覧。このスクリプトが作り直す
      ※ LLM解釈JSON は「{所属クラブ}.json」（ポジションはまとめて1ファイル）を docs/soccer/data/llm/<開幕年>/ に置く。リーグは問わない。

使い方:
    python soccer/scripts/build_soccer_cards.py --data-dir ../sports-dashboard-data/soccer --out docs/soccer/data
    python soccer/scripts/build_soccer_cards.py ... --leagues Premier LaLiga --seasons 2026 --min-minutes 180

カードを作る出場時間の下限:
    既定は、1シーズンを消化したあとは900分。シーズン途中は、消化した試合数×90分×0.3（270〜900分）に下げる
    （run.py の順位の条件と同じ）。--min-minutes で固定できる。

グループ（position_detail）: FW（中央のFW）/ WG（ウイング）/ MF / CB / SB / GK。MFは position_role（DM・CM・AM）で、コア指標が変わる。

依存: pandas numpy openpyxl
"""
import argparse
import hashlib
import json
import re
import shutil
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
MIN_MINUTES_CAP, MIN_MINUTES_SHARE, MIN_MINUTES_FLOOR = 900, 0.30, 270     # run.py の RANK_MIN_* と同じ

# ---- カードに載せる指標の定義 -------------------------------------------------------
# (列名, 表示名, 単位, 良い方向)   ※ 良い方向: high=大きいほど良い / low=小さいほど良い / neutral=良し悪しではない（順位は出さない）
# 列が入力に無い指標（例: Understatを結合していないときのxG・xA、ボール運び）は自動で省く。
# コア指標＝順位つきの主要指標。詳細指標＝テーマ別の補助指標。指標の選び方は「指標整理表」のおすすめ指標に合わせている。
def M(col, label, unit="", better="high"):
    return (col, label, unit, better)


CORE = {
    "FW": [M("np_goals_p90", "ノンPKゴール/90"), M("shots_p90", "シュート/90"),
           M("np_goal_conversion_pct", "ノンPKゴール決定率", "%"), M("big_chances_p90", "ビッグチャンス/90"),
           M("big_chance_conversion_pct", "ビッグチャンス決定率", "%"), M("touches_att_pen_p90", "PA内タッチ/90"),
           M("xg_diff_p90", "xG差分/90")],
    "WG": [M("take_ons_won_p90", "ドリブル成功/90"), M("take_on_pct", "ドリブル成功率", "%"),
           M("key_passes_p90", "キーパス/90"), M("assists_p90", "アシスト/90"),
           M("big_chances_created_p90", "ビッグチャンス創出/90"), M("np_goals_p90", "ノンPKゴール/90"),
           M("progressive_carries_p90", "プログレッシブラン/90")],
    "CB": [M("tackle_win_pct", "タックル成功率", "%"), M("interceptions_p90", "インターセプト/90"),
           M("interceptions_padj_p90", "インターセプト/90（保持率補正）"), M("clearances_p90", "クリア/90"),
           M("aerial_win_pct", "空中戦勝率", "%"), M("dribbled_past_pct", "被ドリブル率", "%", "low")],
    "SB": [M("tackle_win_pct", "タックル成功率", "%"), M("dribbled_past_pct", "被ドリブル率", "%", "low"),
           M("crosses_p90", "クロス/90"), M("key_passes_p90", "キーパス/90"), M("assists_p90", "アシスト/90"),
           M("duel_win_pct", "デュエル勝率", "%"), M("progressive_carries_p90", "プログレッシブラン/90")],
    "GK": [M("gk_save_pct", "セーブ率", "%"), M("goals_conceded_p90", "失点/90", "", "low"),
           M("clean_sheet_pct", "クリーンシート率", "%"), M("pass_pct", "パス成功率", "%"),
           M("long_ball_pct", "ロングボール成功率", "%"), M("gk_sweeper_actions_p90", "エリア外対応/90")],
}
# MFは、役割（position_role）で、コア指標を変える。DM=守備的 / CM=中央 / AM=攻撃的
_MF_BASE = [M("pass_pct", "パス成功率", "%"), M("progressive_passes_p90", "プログレッシブパス/90")]
CORE_MF = {
    "DM": _MF_BASE + [M("tackles_padj_p90", "タックル/90（保持率補正）"), M("interceptions_padj_p90", "インターセプト/90（保持率補正）"),
                      M("tackles_won_p90", "タックル成功/90"), M("interceptions_p90", "インターセプト/90")],
    "CM": _MF_BASE + [M("key_passes_p90", "キーパス/90"), M("tackles_padj_p90", "タックル/90（保持率補正）"),
                      M("interceptions_padj_p90", "インターセプト/90（保持率補正）"), M("ball_recoveries_padj_p90", "ボール奪取/90（保持率補正）")],
    "AM": _MF_BASE + [M("key_passes_p90", "キーパス/90"), M("big_chances_created_p90", "ビッグチャンス創出/90"),
                      M("assists_p90", "アシスト/90"), M("gca_p90", "ゴール創出アクション/90（近似）")],
}
MF_ROLE_LABEL = {"DM": "守備的MF", "CM": "中央MF", "AM": "攻撃的MF"}

_DEF_ZONE_SEC = ("守備", None)
DETAIL = {
    "FW": [("決定力・シュート", [M("goals_p90", "ゴール/90（PK込み）"), M("shot_on_target_pct", "シュート枠内率", "%"),
                             M("shots_in_box_pct", "ボックス内シュート率", "%"), M("headed_shots_p90", "ヘディングシュート/90")]),
           ("チャンス関与", [M("assists_p90", "アシスト/90"), M("gca_p90", "ゴール創出アクション/90（近似）"), M("key_passes_p90", "キーパス/90")]),
           ("空中戦・ポストプレー", [M("aerial_win_pct", "空中戦勝率", "%"), M("fouls_won_p90", "被ファウル/90"),
                                M("take_ons_won_p90", "ドリブル成功/90"), M("dispossessed_p90", "ボールロスト/90", "", "low"),
                                M("bad_touches_p90", "ボールタッチのミス/90", "", "low")]),
           ("守備", [M("def_actions_att_third_p90", "敵陣守備アクション/90"), M("yellow_cards_p90", "イエローカード/90", "", "low")])],
    "WG": [("得点への関与", [M("goals_p90", "ゴール/90（PK込み）"), M("shots_p90", "シュート/90"), M("np_goal_conversion_pct", "ノンPKゴール決定率", "%"),
                        M("touches_att_pen_p90", "PA内タッチ/90"), M("gca_p90", "ゴール創出アクション/90（近似）")]),
           ("クロス・供給", [M("crosses_p90", "クロス/90"), M("cross_pct", "クロス成功率", "%"),
                        M("passes_into_box_p90", "ボックスへのパス/90"), M("progressive_passes_p90", "プログレッシブパス/90")]),
           ("ドリブル・ボール保持", [M("fouls_won_p90", "被ファウル/90"), M("dispossessed_p90", "ボールロスト/90", "", "low"),
                              M("bad_touches_p90", "ボールタッチのミス/90", "", "low")]),
           ("守備", [M("def_actions_att_third_p90", "敵陣守備アクション/90"), M("tackles_won_p90", "タックル成功/90")])],
    "MF": [("パス・ボール保持", [M("passes_p90", "パス/90"), M("forward_pass_pct", "前方パスの割合", "%", "neutral"),
                           M("avg_pass_distance", "平均パス距離", "m", "neutral"), M("dispossessed_p90", "ボールロスト/90", "", "low")]),
           ("攻撃への関与", [M("passes_final_third_p90", "ファイナルサードへのパス/90"), M("passes_into_box_p90", "ボックスへのパス/90"),
                        M("through_balls_p90", "スルーパス/90"), M("shots_p90", "シュート/90"), M("assists_p90", "アシスト/90"),
                        M("big_chances_created_p90", "ビッグチャンス創出/90"), M("gca_p90", "ゴール創出アクション/90（近似）")]),
           ("守備・規律", [M("tackles_won_p90", "タックル成功/90"), M("ball_recoveries_p90", "ボール奪取/90"), M("blocked_passes_p90", "パスブロック/90"),
                        M("duel_win_pct", "デュエル勝率", "%"), M("dribbled_past_p90", "被ドリブル突破/90", "", "low"),
                        M("fouls_committed_p90", "ファウル/90", "", "low"), M("yellow_cards_p90", "イエローカード/90", "", "low")])],
    "CB": [("守備（量・保持率補正）", [M("tackles_padj_p90", "タックル/90（保持率補正）"), M("clearances_padj_p90", "クリア/90（保持率補正）"),
                              M("tackles_won_p90", "タックル成功/90"), M("blocked_passes_p90", "パスブロック/90")]),
           ("空中戦", [M("aerials_won_p90", "空中戦勝利/90")]),
           ("ビルドアップ", [M("pass_pct", "パス成功率", "%"), M("progressive_passes_p90", "プログレッシブパス/90"), M("passes_p90", "パス/90"),
                        M("forward_pass_pct", "前方パスの割合", "%", "neutral"), M("avg_pass_distance", "平均パス距離", "m", "neutral"),
                        M("long_ball_pct", "ロングボール成功率", "%"), M("long_balls_p90", "ロングボール/90")]),
           ("規律・ミス", [M("errors_p90", "エラー/90", "", "low"), M("fouls_committed_p90", "ファウル/90", "", "low"),
                       M("yellow_cards_p90", "イエローカード/90", "", "low"), M("offsides_provoked_p90", "オフサイドを誘った/90")])],
    "SB": [("攻撃参加", [M("big_chances_created_p90", "ビッグチャンス創出/90"), M("cross_pct", "クロス成功率", "%"),
                     M("passes_into_box_p90", "ボックスへのパス/90"), M("progressive_passes_p90", "プログレッシブパス/90"),
                     M("take_ons_won_p90", "ドリブル成功/90")]),
           ("守備・規律", [M("interceptions_p90", "インターセプト/90"), M("tackles_won_p90", "タックル成功/90"),
                       M("tackles_padj_p90", "タックル/90（保持率補正）"), M("fouls_committed_p90", "ファウル/90", "", "low"),
                       M("yellow_cards_p90", "イエローカード/90", "", "low")])],
    "GK": [("被シュート・配球", [M("saves_p90", "セーブ/90", "", "neutral"), M("long_balls_p90", "ロングボール/90", "", "neutral"),
                           M("goal_kick_long_pct", "ロングゴールキックの割合", "%", "neutral"),
                           M("avg_goal_kick_distance", "ゴールキックの平均距離", "m", "neutral"), M("passes_p90", "パス/90", "", "neutral")])],
}

# 列が無いが、ほかの列から作れる指標（順位は付かない）
DERIVED = {
    "shots_in_box_pct": lambda r: (r["shots_in_box"] / r["shots"] * 100) if (r.get("shots") or 0) else None,
    # ボールロスト・ボールタッチのミスは、タッチ数（touches）に対する割合も出す（回数だけだと、出場時間やボール関与量の違いで比較しにくいため）
    "dispossessed_pct": lambda r: (r["dispossessed"] / r["touches"] * 100) if (r.get("touches") or 0) else None,
    "bad_touches_pct": lambda r: (r["bad_touches"] / r["touches"] * 100) if (r.get("touches") or 0) else None,
    # 「成功率」ではなく、全パスのうち前進パスが占める割合（量ではなくスタイルを見る指標）
    "progressive_passes_pct": lambda r: (r["progressive_passes"] / r["passes"] * 100) if (r.get("passes") or 0) else None,
}

# プレータイプ。name はプロンプト付録Bの名前と完全一致させる（LLM解釈JSONの play_evaluations[].name と突き合わせる）。
_ZONES = [("Def 3rd", "自陣3分の1", "def_third"), ("Mid 3rd", "中央3分の1", "mid_third"), ("Att 3rd", "敵陣3分の1", "att_third")]
_DF_TYPES = {
    "count_label": "守備アクション", "share_of": [f"def_actions_{z}" for _, _, z in _ZONES],
    "types": [{"name": n, "label": l, "count": f"def_actions_{z}", "metrics": [
        M(f"tackles_{z}_p90", "タックル/90"), M(f"tackle_win_pct_{z}", "タックル成功率", "%"),
        M(f"interceptions_{z}_p90", "インターセプト/90")]} for n, l, z in _ZONES],
}
_SCA_TYPES = {
    "count_label": "シュートにつながったプレー（近似）", "share_of": "sca", "types": [
        {"name": "Live-ball Pass", "label": "流れの中のパスから", "count": "sca_pass_live", "metrics": [M("sca_pass_live_p90", "回数/90")]},
        {"name": "Dead-ball Pass", "label": "セットプレーから", "count": "sca_pass_dead", "metrics": [M("sca_pass_dead_p90", "回数/90")]},
        {"name": "Take-On", "label": "ドリブル突破から", "count": "sca_take_on", "metrics": [M("sca_take_on_p90", "回数/90")]},
        {"name": "Shot", "label": "こぼれ球・セカンドから", "count": "sca_shot", "metrics": [M("sca_shot_p90", "回数/90")]},
        {"name": "Fouls Drawn", "label": "ファウル奪取から", "count": "sca_foul_won", "metrics": [M("sca_foul_won_p90", "回数/90")]},
        {"name": "Defensive Action", "label": "守備からの転化", "count": "sca_defensive", "metrics": [M("sca_defensive_p90", "回数/90")]},
    ]}
PLAYTYPES = {
    "FW": _SCA_TYPES, "WG": _SCA_TYPES,
    "MF": {"count_label": "パス", "share_of": "passes", "types": [
        {"name": "Short", "label": "短距離", "count": "passes_short", "metrics": [M("pass_short_pct", "成功率", "%")]},
        {"name": "Medium", "label": "中距離", "count": "passes_medium", "metrics": [M("pass_medium_pct", "成功率", "%")]},
        {"name": "Long", "label": "長距離", "count": "passes_long", "metrics": [M("pass_long_pct", "成功率", "%")]},
        {"name": "Dead", "label": "セットプレー", "count": "passes_dead", "metrics": [M("passes_dead_pct", "成功率", "%")]},
        {"name": "FK", "label": "フリーキック", "count": "passes_fk", "metrics": [M("passes_fk_pct", "成功率", "%")]},
        {"name": "TB", "label": "スルーパス", "count": "through_balls", "metrics": [M("through_ball_pct", "成功率", "%")]},
        {"name": "Sw", "label": "サイドチェンジ", "count": "switches", "metrics": [M("switch_pct", "成功率", "%")]},
        {"name": "Crs", "label": "クロス", "count": "crosses", "metrics": [M("cross_pct", "成功率", "%")]},
    ]},
    "CB": _DF_TYPES, "SB": _DF_TYPES,
}

# 試合ログに出す列（共通列のあとに、グループごとの列）
GAMELOG_COMMON = [("minutes", "出場"), ("rating", "評価")]
GAMELOG_EXTRA = {
    "FW": [("goals", "得点"), ("assists", "アシスト"), ("shots", "シュート"), ("shots_on_target", "枠内"), ("key_passes", "キーパス")],
    "WG": [("goals", "得点"), ("assists", "アシスト"), ("shots", "シュート"), ("key_passes", "キーパス"), ("take_ons_won", "ドリブル成功")],
    "MF": [("goals", "得点"), ("assists", "アシスト"), ("passes", "パス"), ("pass_pct", "パス%"), ("key_passes", "キーパス"),
           ("tackles", "タックル"), ("interceptions", "INT")],
    "CB": [("tackles", "タックル"), ("interceptions", "INT"), ("clearances", "クリア"), ("aerials_won", "空中戦勝")],
    "SB": [("assists", "アシスト"), ("tackles", "タックル"), ("interceptions", "INT"), ("crosses", "クロス"), ("key_passes", "キーパス")],
    "GK": [("saves", "セーブ"), ("goals_conceded", "失点"), ("clean_sheets", "無失点"), ("pass_pct", "パス%")],
}
GROUPS = ("FW", "WG", "MF", "CB", "SB", "GK")

# ---- stats：カードに載せる「全指標の生データ」------------------------------------------------
# 画面（player-cards.html）は、この stats から、出したい指標を自分で選んで並べる。
# → どの指標を・どのカテゴリに・どのポジションに出すかを変えるときは、HTML だけを直せばよく、このスクリプトの再実行は要らない。
# 形式: {列名: [全体, ホーム, アウェイ]}、各要素は [値, 順位, 母数]（順位が無ければ [値]、その試合区分に値が無ければ null）。
# run.py に順位が無い指標は、[値, 大きい順の順位, 母数, 小さい順の順位] の4要素（add_derived_ranks）
STATS_SKIP = {"player_id", "touch_pos_n", "pass_distance", "goal_kick_distance",          # ID・内部用の合計値は載せない
              "touch_x_sum", "touch_y_sum", "touch_width_sum"}
STATS_SKIP_SUFFIX = ("_順位", "_順位_母数")

# 一覧の並び替えに使う値（venue=all の行から取る）
INDEX_SORT_COLS = ["rating_avg", "minutes", "matches", "goals_p90", "assists_p90", "np_goals_p90", "key_passes_p90",
                   "big_chances_created_p90", "progressive_passes_p90", "tackles_p90", "interceptions_p90"]


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


def min_minutes_for(players, fixed):
    """カードを作る出場時間の下限。fixed が無ければ、消化した試合数に合わせる（run.py の順位の条件と同じ）。"""
    if fixed is not None:
        return float(fixed)
    games = pd.to_numeric(players.loc[players["venue"] == "all", "matches"], errors="coerce").max()
    if not games or np.isnan(games):
        return float(MIN_MINUTES_CAP)
    return float(min(MIN_MINUTES_CAP, max(MIN_MINUTES_FLOOR, MIN_MINUTES_SHARE * games * 90)))


# ---- カードの部品 --------------------------------------------------------------------
def venue_rows(g):
    """選手×クラブの、all/home/away の3行を {venue: Series} にする。"""
    return {v: r for v, r in ((row["venue"], row) for _, row in g.iterrows()) if v in VENUES}


def cell(rows, venue, col, better):
    r = rows.get(venue)
    if r is None:
        return None
    if col in r.index:
        v = num(r[col])
    elif col in DERIVED:
        v = num(DERIVED[col](r))
    else:
        return None
    if v is None:
        return None
    if better == "neutral":                          # 良し悪しではない指標には、順位を付けない
        return {"v": v, "rank": None, "pop": None}
    return {"v": v, "rank": num(r.get(col + "_順位")), "pop": num(r.get(col + "_順位_母数"))}


def metric_entry(rows, spec, columns):
    col, label, unit, better = spec
    if col not in columns and col not in DERIVED:
        return None
    entry = {"key": col, "label": label, "unit": unit, "better": better}
    for v in VENUES:
        entry[v] = cell(rows, v, col, better)
    return entry if any(entry[v] for v in VENUES) else None


def stats_entries(rows, stat_cols):
    """選手×クラブの全数値列を、コンパクトな形（stats）にする。累計は列 xxx、90分あたりは列 xxx_p90 として入る。"""
    out = {}
    for col in stat_cols:
        cells = []
        for v in VENUES:
            r = rows.get(v)
            val = None
            if r is not None:
                val = num(r[col]) if col in r.index else num(DERIVED[col](r))
            if val is None:
                cells.append(None)
                continue
            rk, pop = (num(r.get(col + "_順位")), num(r.get(col + "_順位_母数"))) if r is not None else (None, None)
            cells.append([val, rk, pop] if rk is not None and pop is not None else [val])
        if any(c is not None for c in cells):
            out[col] = cells
    return out


# run.py が順位（_順位）を作っていない指標にも、順位を付ける（画面で、すべての指標に順位とカラースケールを出すため）。
# 対象：90分あたり（_p90）・割合（_pct）・平均（avg_）と、下の EXTRA。累計の回数は対象外。
# 同じリーグ・シーズン・ポジション（FW/WG/MF/CB/SB/GK）で、カードを作った選手の中での順位。
# 値の大きい順（rank_desc）と小さい順（rank_asc）の両方を入れ、どちらを使うか（大きいほど良い/小さいほど良い）は画面側で決める。
# run.py が一部の選手にだけ順位を付けている指標（試行数の下限があるもの）は、その基準を尊重して、ここでは順位を足さない。
RANK_EXTRA = {"red_cards", "carries", "carries_final_third", "carries_into_box", "progressive_carries"}


def wants_derived_rank(col):
    return col.endswith(("_p90", "_pct")) or col.startswith("avg_") or col in RANK_EXTRA


def add_derived_ranks(cards):
    by_group = {}
    for c in cards:
        by_group.setdefault(c["group"], []).append(c)
    for cs in by_group.values():
        cols = {k for c in cs for k in c["stats"] if wants_derived_rank(k)}
        for col in cols:
            for vi in range(len(VENUES)):
                cells = [(c, c["stats"][col][vi]) for c in cs if col in c["stats"] and c["stats"][col][vi] is not None]
                if not cells or any(len(cell) >= 3 for _, cell in cells):
                    continue                                       # 順位が無い、または run.py が順位を付けている指標
                vals = pd.Series([cell[0] for _, cell in cells], dtype=float)
                desc, asc = vals.rank(ascending=False, method="min"), vals.rank(ascending=True, method="min")
                for (c, cell), d, a in zip(cells, desc, asc):
                    c["stats"][col][vi] = [cell[0], int(d), len(vals), int(a)]      # [値, 大きい順の順位, 母数, 小さい順の順位]


def core_specs(group, role):
    if group == "MF":
        return CORE_MF.get(role) or CORE_MF["CM"]
    return CORE[group]


def detail_entries(group, rows, columns, core_keys):
    """テーマ別の詳細指標。コア指標と重なるものは省く。"""
    out = []
    for title, specs in DETAIL.get(group, []):
        ms = [m for m in (metric_entry(rows, s, columns) for s in specs if s[0] not in core_keys) if m]
        if ms:
            out.append({"title": title, "metrics": ms})
    return out


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


def build_cards(players, matches, league, year, min_minutes_fixed):
    columns = set(players.columns)
    threshold = min_minutes_for(players, min_minutes_fixed)
    games = num(pd.to_numeric(players.loc[players["venue"] == "all", "matches"], errors="coerce").max())   # 消化した試合数の目安
    players = players[players["position"].notna()].copy()
    players["group"] = np.where(players["position"] == "DF", players["position_detail"], players["position"])
    stat_cols = [c for c in players.columns if pd.api.types.is_numeric_dtype(players[c]) and c not in STATS_SKIP
                 and not c.endswith(STATS_SKIP_SUFFIX)] + [c for c in DERIVED if c not in players.columns]
    mkey = {k: g for k, g in matches.groupby(["team", "player_id"])}
    cards, index = [], []
    for (team, pid), g in players.groupby(["team", "player_id"]):
        rows = venue_rows(g)
        a = rows.get("all")
        if a is None or (num(a["minutes"]) or 0) < threshold:
            continue
        group = a["group"]
        if group not in GROUPS:
            continue
        role = text(a.get("position_role")) if "position_role" in columns else None
        cid = f"{int(pid)}_{slug(team)}"
        log_cols, log = game_log(mkey[(team, pid)], group) if (team, pid) in mkey else ([], [])
        specs = core_specs(group, role)
        core = [m for m in (metric_entry(rows, s, columns) for s in specs) if m]
        card = {
            "id": cid, "name": a["player"], "team": team, "league": league, "season": str(year),
            "position": a["position"], "group": group, "role": role,
            "matches": num(a["matches"]), "starts": num(a["starts"]), "minutes": num(a["minutes"]),
            "rating_avg": num(a.get("rating_avg")), "motm": num(a.get("motm")),
            "goals": num(a.get("goals")), "assists": num(a.get("assists")), "clean_sheets": num(a.get("clean_sheets")),
            "stats": stats_entries(rows, stat_cols),          # 画面が使う全指標（表示の選び方は HTML 側で決める）
            "core": core,                                     # 旧形式（互換用。画面は stats があれば使わない）
            "detail": detail_entries(group, rows, columns, {m["key"] for m in core}),
            "playtypes": playtype_entries(group, rows, columns),
            "game_log_cols": log_cols, "game_log": log,
        }
        cards.append(card)
        idx = {"id": cid, "name": a["player"], "team": team, "position": a["position"], "group": group, "role": role}
        for c in INDEX_SORT_COLS:
            idx[c] = num(a.get(c)) if c in columns else None
        index.append(idx)
    add_derived_ranks(cards)
    return cards, index, threshold, games


# ---- 出力 --------------------------------------------------------------------------
def write_league_season(out_dir, league, year, cards, index, games=None):
    numeric = out_dir / "games" / str(year) / league
    numeric.mkdir(parents=True, exist_ok=True)
    for old in numeric.glob("*.json"):          # 閾値の変更や移籍で消えた選手のファイルを残さない
        old.unlink()
    for c in cards:
        write_json(numeric / f"{c['id']}.json", c)
    index.sort(key=lambda p: (p["team"], p["name"]))
    write_json(numeric / "index.json", {
        "league": league, "season": str(year), "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "games": games, "teams": sorted({p["team"] for p in index}), "players": index})
    return len(cards)


def write_llm_indexes(out_dir):
    """LLM解釈JSON（手動で置く）は data/llm/<年>/<クラブ名>.json。年ごとの index.json をここで作り直す。"""
    for d in sorted((out_dir / "llm").glob("*")):
        if d.is_dir():
            files = sorted(f.name for f in d.glob("*.json") if f.name != "index.json")
            write_json(d / "index.json", {"files": files})


def cleanup_legacy(out_dir):
    """旧構成（<League>/<年>/cards_numeric）の自動生成フォルダを消す。"""
    for lg in list(LEAGUE_LABELS):
        d = out_dir / lg
        if d.is_dir() and any(d.glob("*/cards_numeric")):
            shutil.rmtree(d)
            print(f"旧フォルダを削除: {d}")


def write_manifest(out_dir):
    """出力フォルダにある games/<年>/<League>/index.json を走査して作るので、一部のリーグだけ更新しても他が消えない。
    llm_batches = そのリーグ・シーズンのクラブのうち、llm/<年>/ にJSONがあるクラブ数。
    games = そのリーグ・シーズンで消化した試合数の目安（画面が、試合数の少ない年度を初期表示にしないために使う）。"""
    leagues = {}
    for idx in sorted(out_dir.glob("games/*/*/index.json")):
        year, league = idx.parts[-3], idx.parts[-2]
        meta = json.loads(idx.read_text(encoding="utf-8"))
        teams = set(meta.get("teams", []))
        lidx = out_dir / "llm" / year / "index.json"
        have = json.loads(lidx.read_text(encoding="utf-8")).get("files", []) if lidx.exists() else []
        n_llm = sum(1 for f in have if f[:-5] in teams)
        leagues.setdefault(league, {"label": LEAGUE_LABELS.get(league, league), "seasons": []})["seasons"].append(
            {"year": year, "label": season_label(int(year)), "llm_batches": n_llm, "games": meta.get("games")})
    for v in leagues.values():
        v["seasons"].sort(key=lambda s: s["year"], reverse=True)
    order = [k for k in LEAGUE_LABELS if k in leagues] + [k for k in leagues if k not in LEAGUE_LABELS]
    write_json(out_dir / "manifest.json", {"leagues": {k: leagues[k] for k in order}})


FILE_RE = re.compile(r"all_leagues_(players|matches)_(\d{4})(?:_([A-Za-z0-9]+))?\.xlsx")


def find_sources(data_dir, seasons):
    """
    {開幕年: {"players": [ファイル…], "matches": [ファイル…]}}。
    soccer/2025年/all_leagues_players_2025.xlsx など。フォルダ名は問わず、ファイル名の年で判断する（1ファイル1シート）。
    リーグ別のファイル（all_leagues_matches_2025_Premier.xlsx）も拾う。全リーグ1ファイルが複数あるときは、新しい方だけを使う。
    """
    src = {}
    for f in sorted(data_dir.rglob("all_leagues_*_*.xlsx")):
        m = FILE_RE.fullmatch(f.name)
        if not m or (seasons and int(m.group(2)) not in seasons):
            continue
        year, kind, part = int(m.group(2)), m.group(1), m.group(3)
        src.setdefault(year, {}).setdefault(kind, {"whole": [], "parts": []})["parts" if part else "whole"].append(f)
    out = {}
    for year, kinds in src.items():
        out[year] = {}
        for kind, d in kinds.items():
            whole = sorted(d["whole"], key=lambda f: f.stat().st_mtime)
            if len(whole) > 1:
                print(f"注意: {year} の {kind} が複数あります。新しい方を使います: {whole[-1]}（他: {[str(f) for f in whole[:-1]]}）")
            out[year][kind] = {"whole": whole[-1:], "parts": d["parts"]}
    return out


def load_kind(files, leagues=None):
    """リーグ別のファイルを先に読み、全リーグ1ファイルからは、リーグ別のファイルがあるリーグの分を除いて足す。"""
    frames, covered = [], set()
    for f in files["parts"]:
        df = pd.read_excel(f)
        covered |= set(df["league"])
        frames.append(df)
    for f in files["whole"]:
        df = pd.read_excel(f)
        frames.append(df[~df["league"].isin(covered)])
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data-dir", required=True, type=Path,
                    help="年ごとのフォルダ（2025年/ など）が入っているフォルダ（soccer）。中を再帰的に探す")
    ap.add_argument("--out", required=True, type=Path, help="出力先（docs/soccer/data）")
    ap.add_argument("--leagues", nargs="*", help="対象リーグ（省略で全部）")
    ap.add_argument("--seasons", nargs="*", type=int, help="対象の開幕年（省略で全部）")
    ap.add_argument("--min-minutes", type=float, default=None,
                    help="カードを作る出場時間の下限（分）。省略時は、1シーズンを消化したあとは900分、シーズン途中は消化した試合数に合わせて下げる")
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
        players = load_kind(s["players"])
        matches = load_kind(s["matches"])
        for league in [l for l in LEAGUE_LABELS if l in set(players["league"])] + \
                      sorted(set(players["league"]) - set(LEAGUE_LABELS)):
            if args.leagues and league not in args.leagues:
                continue
            cards, index, thr, games = build_cards(players[players["league"] == league], matches[matches["league"] == league],
                                                   league, year, args.min_minutes)
            if not cards:
                print(f"[{league} {year}] 出場{thr:.0f}分以上の選手がいないため、スキップ")
                continue
            n = write_league_season(args.out, league, year, cards, index, games)
            print(f"[{league} {year}] カード{n}枚（出場{thr:.0f}分以上・{games}試合消化）→ {args.out / 'games' / str(year) / league}")
    cleanup_legacy(args.out)
    write_llm_indexes(args.out)
    write_manifest(args.out)


if __name__ == "__main__":
    main()
