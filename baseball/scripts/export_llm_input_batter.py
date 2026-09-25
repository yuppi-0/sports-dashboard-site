# -*- coding: utf-8 -*-
"""
export_llm_input_batter.py
====================
投手版 export_llm_input.py と対になる打者版。run.py の run_dashboard() が
生成している「日別ダッシュボードJSON」（games/json/{date}.json）を読み込み、
打者のシーズン集計・対左右投手別成績を算出して、
  1) LLM入力用xlsx（縦持ち・複数シート）
  2) batter-cards.html が読み込む数値データJSON（batter_cards_numeric/配下）
を出力する。

投手版との違い・制約:
  - 対右投手/対左投手のスプリットは、球種別集計（pitchSplits経由）がある試合では
    球単位の実際の対戦投手の利き腕で判定される。pitchSplitsが無い古いデータ
    （移行前に生成された日別JSON）では、従来通り「その試合の相手チーム先発投手の
    投球腕」による近似にフォールバックする。
  - 球種別・球速帯別の内訳（byPitchType）は、run.py（NPB）・run_mlb.py（MLB）が
    日別JSONの各打者エントリに付与する"pitchSplits"（試合×球種×球速帯×対戦投手
    利き腕、で事前集計済み）を元に、シーズン全体で合算して組み立てる。球速帯は
    球種コードごとの固定ビン（export_llm_input_batter.pyのPITCH_VELO_BANDS。
    詳細はbatter_card_schema.md参照）。この内訳のOBP/SLG/OPSは、単打・長打の
    内訳が無く計算できないため含めていない（打率・選球眼系のみ）。
  - 出塁率/長打率/OPS（シーズン全体・対左右投手）は、単打・二塁打・三塁打の内訳が
    日別JSON側に無く総塁打数を正確に積み上げられないため、試合ごとに計算済みの値を
    打席数で加重平均した近似値。安打・本塁打・四死球・三振・打点・盗塁・打率・K%・BB%
    は実数の積み上げなので正確。
  - 守備・走塁の高度指標（守備率・レンジファクター・盗塁死・盗塁成功率など）は、
    現時点のrun.pyに守備成績ページのスクレイピングが実装されていないため取得できない。
    numeric_json側にはキーだけ用意してNoneを入れてある。追加するには、run.py側に
    守備成績（刺殺・補殺・失策・捕逸）と盗塁死を取得する新しいスクレイピングの
    ステップを追加する必要がある（別途対応）。
  - run.py（NPB）・run_mlb.py（MLB）どちらの games/json からも同じ関数で処理できるよう
    共通の入力形式（games/json/{date}.json、pitchers.hand・batters.pa等）に依存する
    作りにしている。ただしMLB側のbatterエントリには、NPB側にある試合単位の
    chase(O-Swing%)/whiff(whiff%)/contact(Z-Swing%) の3項目が無い（Statcastベースの
    Hard-Hit%/Barrel%/xwOBA等に置き換わっている）ため、season（試合単位の値を加重平均
    する既存ロジック）の"O-Swing%"/"Z-Swing%"/"whiff%"はMLBでは必ずNoneになる。
    一方、pitchSplits経由のbyPitchType（球種別・球速帯別の内訳）側は、run_mlb.pyの
    build_batter_pitch_splits_mlb()がStatcastの description/zone から直接計算するため、
    MLBでも欠損しない（chase_pct/contact_pct/whiff_pctが球種ごとに取れる）。
  - MLB側の"bb"は「四球」のみ（死球を含まない）。NPB側の"bb"は「四死球」（四球+死球）。
    このスクリプトはどちらも同じ"bb"キーとして合算するため、リーグ間で出塁率近似の
    厳密な計算根拠が微妙に異なる点に注意（どちらも近似値であることに変わりはない）。

使い方:
    python export_llm_input_batter.py \
        --games-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/games/json" \
        --out "data/datamart/llm_input/プロ野球_2026_打者データ.xlsx" \
        --numeric-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/batter_cards_numeric" \
        --min-pa 0
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
from collections import Counter

import pandas as pd

# 日別JSONの読み込みロジックは投手版と完全に共通のため使い回す
from export_llm_input import load_daily_games


# ==================================================
# Section 0. 球速帯の固定ビン定義（run.py・run_mlb.py共有）
# ==================================================
# 球種コード別の球速帯を「遅い/中間/速い」の3分割で判定する（旧版は球種ごとに
# 4分割の実数レンジ(例:150-154)だったが、見づらいため3分割の相対ラベルに統一）。
# ロジック: 各球種コードにつき2つの閾値(lo, hi)を持ち、
#   velo < lo        → 遅い
#   lo <= velo < hi  → 中間
#   velo >= hi       → 速い
# 閾値は球種ごとの典型的な球速分布を踏まえた固定値（球種によって球速レンジの傾向は
# 概ね決まっているため、球種横断の固定ビンではなく球種コードごとに用意する。実データを
# 見ながら調整可能な定数）。NPB(run.py)・MLB(run_mlb.py)どちらの打者版球種集計も
# ここから同じ定義を import して使う（選手間・リーグ間比較のため、境界は揃えておく必要がある）。
PITCH_VELO_BANDS = {
    "FF": (145, 155),   # ストレート/フォーシーム
    "SI": (140, 150),   # シンカー
    "CT": (135, 145),   # カットボール
    "SH": (135, 145),   # シュート
    "SL": (125, 135),   # スライダー
    "CU": (110, 120),   # カーブ
    "FK": (115, 125),   # フォーク
    "FS": (120, 130),   # スプリット
}
PITCH_VELO_BANDS_DEFAULT = (120, 140)

# 球速帯の表示ラベル（この順で「遅い→中間→速い」。batter-cards.html側のソート順にも使う）
VELO_BAND_LABELS = ("遅い", "中間", "速い")


def velo_band_label(pitch_code: str, velo) -> str | None:
    """球種コードと球速(km/h)から3分割の球速帯ラベル（遅い/中間/速い）を返す。
    球速が無ければNone。"""
    if velo is None or (isinstance(velo, float) and pd.isna(velo)):
        return None
    lo, hi = PITCH_VELO_BANDS.get(pitch_code, PITCH_VELO_BANDS_DEFAULT)
    if velo < lo:
        return VELO_BAND_LABELS[0]
    if velo < hi:
        return VELO_BAND_LABELS[1]
    return VELO_BAND_LABELS[2]


# ==================================================
# Section 0b. 球種カテゴリ階層（球種詳細 → 球種中カテゴリ → 球種大カテゴリ）
# ==================================================
# 球種詳細（pitchSplitsの"t"にそのまま入っている表記ゆれ込みの球種名）から、
# 中カテゴリ（ストレート系/スライダー系/カーブ系/フォーク系/シンカー系/シュート系）、
# さらに大カテゴリ（ストレート系/曲がる系/落ちる系）を引けるようにするマッピング。
# 大カテゴリの分類方針: ストレート系＝速球そのもの／曲がる系＝横方向の変化が主体
# （スライダー・シュート・シンカー）／落ちる系＝縦方向の変化が主体（フォーク・カーブ）。
PITCH_MID_CATEGORY = {
    "ストレート": "ストレート系", "フォーシーム": "ストレート系",
    "ツーシーム": "ストレート系", "ワンシーム": "ストレート系",
    "スライダー": "スライダー系", "カットボール": "スライダー系",
    "スイーパー": "スライダー系", "縦スライダー": "スライダー系", "スラーブ": "スライダー系",
    "カーブ": "カーブ系", "ナックルカーブ": "カーブ系", "スローボール": "カーブ系", "スローカーブ": "カーブ系",
    "フォーク": "フォーク系", "チェンジアップ": "フォーク系", "スプリット": "フォーク系",
    # シンカー/シュートは動きの系統が同じ（腕側に沈む球）ため、中カテゴリは1つにまとめる
    "シンカー": "シンカー系",
    "シュート": "シンカー系",
}
PITCH_MAJOR_CATEGORY = {
    "ストレート系": "ストレート系",
    "スライダー系": "曲がる系", "シンカー系": "曲がる系",
    "カーブ系": "落ちる系", "フォーク系": "落ちる系",
}

# 球種詳細・中カテゴリの表示順（PITCH_MID_CATEGORYを定義した並び＝中カテゴリごとに
# まとめた順が、そのまま球種詳細の表示順になる。batter-cards.html側のプルダウン・
# 軸の値一覧もbyPitchTypeの並び順をそのまま使うため、ここで一度だけ順序を決めておけば
# バックエンド・フロントエンドで表示順がずれない）。
PITCH_DETAIL_ORDER = list(PITCH_MID_CATEGORY.keys())
PITCH_MID_ORDER = ["ストレート系", "スライダー系", "カーブ系", "フォーク系", "シンカー系"]
PITCH_MAJOR_ORDER = ["ストレート系", "曲がる系", "落ちる系"]


def pitch_category(pitch_type: str) -> tuple[str, str]:
    """球種詳細名から (球種中カテゴリ, 球種大カテゴリ) を返す。未知の球種は「その他」扱い。"""
    mid = PITCH_MID_CATEGORY.get(pitch_type, "その他")
    major = PITCH_MAJOR_CATEGORY.get(mid, "その他")
    return mid, major


def _pitch_sort_key(pitch_type: str) -> tuple[int, int]:
    """球種詳細を「中カテゴリの表示順→中カテゴリ内の表示順」で並べるためのソートキー。
    未知の球種（マッピングに無い＝「その他」）は末尾に回す。"""
    mid, _ = pitch_category(pitch_type)
    mi = PITCH_MID_ORDER.index(mid) if mid in PITCH_MID_ORDER else len(PITCH_MID_ORDER)
    di = PITCH_DETAIL_ORDER.index(pitch_type) if pitch_type in PITCH_DETAIL_ORDER else len(PITCH_DETAIL_ORDER)
    return (mi, di)


# ==================================================
# Section 0c. 守備位置の表記統一
# ==================================================
# run.py/run_mlb.pyが吐く守備位置は "(三)" のようにカッコ付きの場合がある
# （元データの表記ゆれをそのまま引き継いでいる）。batter-cards.html側の表示・
# フィルタでは "三" のようにカッコ無しで統一して使うため、ここで一箇所に集約して剥がす。
def _clean_pos(pos) -> str | None:
    if not pos:
        return pos
    return re.sub(r"[（）()]", "", str(pos)).strip() or None


# コースゾーンの5x5(25分割)グリッド境界（run.py・run_mlb.py共有）。
# 投手側の座標ヒートマップ（pitcher-cards.htmlのhmBuildGridHtml/edges25）と同じ境界。
# ZONE=1.0がストライクゾーンの端、OUTER=1.67がそこからさらに外側の「際どいボール球」の端。
# NPB（run.py）はコース(Left)/コース(Top)から、MLB（run_mlb.py）はplate_x/plate_zから、
# それぞれこの共通の[-1,1]正規化空間に変換したcx/cyをこのグリッドで判定する。
COURSE_ZONE_EDGES = [-1.67, -1.0, -1.0 / 3, 1.0 / 3, 1.0, 1.67]


def hm_get_cell(v: float, edges: list = COURSE_ZONE_EDGES) -> int:
    """pitcher-cards.htmlのhmGetCell()と同じ二分探索無しの線形版（Python移植）"""
    for i in range(len(edges) - 1):
        if edges[i] <= v < edges[i + 1]:
            return i
    return 0 if v < edges[0] else len(edges) - 2


# ==================================================
# Section 1. 打者名収集・appearances構築
# ==================================================

def _normalize_name(name):
    """表記ゆれ吸収: 'Last, First' 形式を 'First Last' 形式に統一する（投手版と同じロジック）"""
    if not isinstance(name, str):
        return name
    name = name.strip()
    if "," in name:
        parts = name.split(",", 1)
        last, first = parts[0].strip(), parts[1].strip()
        if last and first:
            return f"{first} {last}"
    return name


def build_all_batter_names(all_data: dict) -> set[str]:
    """全登場打者名を収集（表記ゆれを正規化・壊れたエントリはスキップ）"""
    names = set()
    for date, games in all_data.items():
        if date == "highlights" or date.startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict):
                continue
            batters = g.get("batters")
            if not isinstance(batters, dict):
                continue
            for side in ("home", "away"):
                for b in (batters.get(side) or []):
                    if isinstance(b, dict) and b.get("name"):
                        names.add(_normalize_name(b["name"]))
    return names


def _opp_pitcher_hand(game: dict, batter_side: str) -> str | None:
    """打者側(home/away)から見た相手チーム先発投手の投球腕(R/L)を返す。
    role/役割 == "先発" のエントリを優先的に探し、無ければ（役割情報が無いデータ側の
    保険として）先頭要素を先発とみなす。取れなければNone。"""
    pitchers = game.get("pitchers") or {}
    if not isinstance(pitchers, dict):
        return None
    opp_side = "away" if batter_side == "home" else "home"
    opp_list = pitchers.get(opp_side) or []
    opp_list = [p for p in opp_list if isinstance(p, dict)]
    if not opp_list:
        return None
    starter = next((p for p in opp_list if p.get("role") == "先発"), opp_list[0])
    return starter.get("hand")


def build_appearances_batter(all_data: dict, player_name: str) -> list[dict]:
    """特定打者の全出場試合を集める。各要素に対戦相手投手の腕(opp_pitcher_hand)を添える"""
    appearances = []
    dates = sorted(d for d in all_data.keys() if d != "highlights" and not d.startswith("_"))
    name_norm = _normalize_name(player_name).lower()
    for date in dates:
        for g in all_data.get(date, []):
            if not isinstance(g, dict):
                continue
            batters = g.get("batters")
            if not isinstance(batters, dict):
                continue
            for side in ("home", "away"):
                for b in (batters.get(side) or []):
                    if not isinstance(b, dict):
                        continue
                    raw_name = b.get("name") or ""
                    if _normalize_name(raw_name).lower() == name_norm:
                        appearances.append({
                            "date": date, "game": g, "side": side, "player": b,
                            "opp_pitcher_hand": _opp_pitcher_hand(g, side),
                        })
    return appearances


# ==================================================
# Section 2. シーズン集計
# ==================================================

def _weighted_avg(pairs) -> float | None:
    """[(value, weight), ...] の加重平均。重み合計が0/valueが全てNoneならNone"""
    num = sum(v * w for v, w in pairs if v is not None and w)
    den = sum(w for v, w in pairs if v is not None and w)
    return round(num / den, 3) if den > 0 else None


def calc_season_batter_stats(appearances: list[dict], hand_filter: str | None = None) -> dict:
    """
    シーズン（または対右/対左投手）打撃成績を1行返す。
    hand_filter: None=全体, "R"=対右投手, "L"=対左投手
    """
    rows = [ap for ap in appearances if isinstance(ap.get("player"), dict)]
    if hand_filter:
        rows = [ap for ap in rows if ap.get("opp_pitcher_hand") == hand_filter]

    if not rows:
        return {
            "試合数": 0, "打席": 0, "打数": 0, "安打": 0, "本塁打": 0, "四死球": 0,
            "三振": 0, "打点": 0, "盗塁": 0, "盗塁死": None, "打率": None, "出塁率": None, "長打率": None,
            "OPS": None, "K%": None, "BB%": None,
            "O-Swing%": None, "Z-Swing%": None, "whiff%": None,
        }

    def s(key):
        return sum((ap["player"].get(key) or 0) for ap in rows)

    pa, ab, h, hr = s("pa"), s("ab"), s("h"), s("hr")
    bb, k, rbi, sb = s("bb"), s("k"), s("rbi"), s("sb")
    # 盗塁死(cs)はMLBのみ取得できる（run_mlb.pyが走者を正しく特定して算出）。
    # NPBの打者エントリには"cs"キー自体が無いため、その場合はNoneのままにしておく
    # （0と区別する。0は「盗塁死が無かった」、Noneは「そもそもデータが無い」）。
    has_cs = any("cs" in (ap["player"] or {}) for ap in rows)
    cs = s("cs") if has_cs else None

    avg = round(h / ab, 3) if ab > 0 else None
    kpct = round(k / pa * 100, 1) if pa > 0 else None
    bbpct = round(bb / pa * 100, 1) if pa > 0 else None

    # OBP/SLG/OPSは打席内訳が無いため、試合ごとの計算済み値を打席数で加重平均した近似値
    obp = _weighted_avg([(ap["player"].get("obp"), ap["player"].get("pa") or 0) for ap in rows])
    slg = _weighted_avg([(ap["player"].get("slg"), ap["player"].get("pa") or 0) for ap in rows])
    ops = _weighted_avg([(ap["player"].get("ops"), ap["player"].get("pa") or 0) for ap in rows])

    chase = _weighted_avg([(ap["player"].get("chase"), ap["player"].get("pa") or 0) for ap in rows])
    contact = _weighted_avg([(ap["player"].get("contact"), ap["player"].get("pa") or 0) for ap in rows])
    whiff = _weighted_avg([(ap["player"].get("whiff"), ap["player"].get("pa") or 0) for ap in rows])

    return {
        "試合数": len(rows), "打席": pa, "打数": ab, "安打": h, "本塁打": hr,
        "四死球": bb, "三振": k, "打点": rbi, "盗塁": sb, "盗塁死": cs,
        "打率": avg, "出塁率": obp, "長打率": slg, "OPS": ops,
        "K%": kpct, "BB%": bbpct,
        "O-Swing%": chase, "Z-Swing%": contact, "whiff%": whiff,
    }


def determine_primary_position(appearances: list[dict]) -> str | None:
    """出場試合の守備位置の最頻値を、その打者の主定位置とみなす"""
    positions = [ap["player"].get("pos") for ap in appearances if isinstance(ap.get("player"), dict)]
    positions = [p for p in positions if p]
    if not positions:
        return None
    return Counter(positions).most_common(1)[0][0]


# ==================================================
# Section 2b. 球種別・球速帯別集計（対左右投手の3系統: all/vsR/vsL）
# ==================================================
# 元データは run.py（NPB）/ run_mlb.py（MLB）が日別JSONの各打者エントリに付与する
# "pitchSplits"（試合×球種×球速帯×対戦投手利き腕、で事前集計済みのリスト）。
# どちらのリーグも同じキー構成（t/band/h/n/pa/ab/h_/2b/3b/hr/bb/hbp/k/sw/ws/z/oz/zsw/ozsw）
# で出力するため、ここのロジックはリーグに依存しない。ただし"rbi"（打点）はMLBの
# pitchSplits等にしか無い（NPBの投球チャートデータには得点状況の情報が無く、
# コース別・球種別・カウント別・状況別の粒度では打点を算出できないため）。
# 2b/3b/hr/bb/hbp/k/rbiは、この内訳を追加した時点（2026年版）以降に生成された
# データにしか無い。無い場合は.get(...,0)で0扱いになり、総塁打・出塁率の分母が0に
# なるので、None（"-"表示）に自然にフォールバックする（古いキャッシュ済みJSONを
# 読んでも壊れない）。rbiは「entriesのどれか1つでもrbiキーを持っていればMLBのデータ、
# 1つも持っていなければNPBか旧データ」とみなしNoneのままにする（0との混同を避ける）。

def _agg_pitch_group(entries: list[dict]) -> dict:
    """pitchSplitsエントリのリストから、打率・長打率・出塁率・OPS・本塁打・打点・
    K%・BB%・選球眼指標を集計した1オブジェクトを作る。"""
    n = sum(e.get("n", 0) or 0 for e in entries)
    pa = sum(e.get("pa", 0) or 0 for e in entries)
    ab = sum(e.get("ab", 0) or 0 for e in entries)
    h = sum(e.get("h_", 0) or 0 for e in entries)
    doubles = sum(e.get("2b", 0) or 0 for e in entries)
    triples = sum(e.get("3b", 0) or 0 for e in entries)
    hr = sum(e.get("hr", 0) or 0 for e in entries)
    bb = sum(e.get("bb", 0) or 0 for e in entries)
    hbp = sum(e.get("hbp", 0) or 0 for e in entries)
    k = sum(e.get("k", 0) or 0 for e in entries)
    has_rbi = any("rbi" in e for e in entries)
    rbi = sum(e.get("rbi", 0) or 0 for e in entries) if has_rbi else None
    sw = sum(e.get("sw", 0) or 0 for e in entries)
    ws = sum(e.get("ws", 0) or 0 for e in entries)
    z = sum(e.get("z", 0) or 0 for e in entries)
    oz = sum(e.get("oz", 0) or 0 for e in entries)
    zsw = sum(e.get("zsw", 0) or 0 for e in entries)
    ozsw = sum(e.get("ozsw", 0) or 0 for e in entries)

    singles = max(h - doubles - triples - hr, 0)
    total_bases = singles + doubles * 2 + triples * 3 + hr * 4
    obp_den = ab + bb + hbp  # 犠飛は元データに無いため近似（他のOBP計算と同じ扱い）

    return {
        "count": n, "pa": pa, "ab": ab, "h": h, "hr": hr, "rbi": rbi,
        "avg": round(h / ab, 3) if ab > 0 else None,
        "slg": round(total_bases / ab, 3) if ab > 0 else None,
        "obp": round((h + bb + hbp) / obp_den, 3) if obp_den > 0 else None,
        "ops": (round(total_bases / ab, 3) + round((h + bb + hbp) / obp_den, 3))
               if ab > 0 and obp_den > 0 else None,
        "k_pct": round(k / pa * 100, 1) if pa > 0 else None,
        "bb_pct": round(bb / pa * 100, 1) if pa > 0 else None,
        "whiff_pct": round(ws / sw * 100, 1) if sw > 0 else None,
        "chase_pct": round(ozsw / oz * 100, 1) if oz > 0 else None,     # O-Swing%
        "contact_pct": round(zsw / z * 100, 1) if z > 0 else None,       # Z-Swing%
    }


def build_by_pitch_type(appearances: list[dict], hand_filter: str | None = None) -> list[dict]:
    """
    appearancesからbyPitchType（all/vsR/vsLのうち1系統ぶん）を組み立てる。
    各球種の中に、さらにbyVelocityBand（球速帯別の内訳）をネストして持たせる
    （batter_card_schema.md参照）。
    """
    entries: list[dict] = []
    for ap in appearances:
        p = ap.get("player")
        if not isinstance(p, dict):
            continue
        for e in (p.get("pitchSplits") or []):
            if hand_filter and e.get("h") != hand_filter:
                continue
            entries.append(e)

    by_type: dict[str, list[dict]] = {}
    for e in entries:
        by_type.setdefault(e.get("t", ""), []).append(e)

    result = []
    for pitch_type, type_entries in by_type.items():
        stats = _agg_pitch_group(type_entries)
        stats["pitchType"] = pitch_type
        stats["pitchCategoryMid"], stats["pitchCategoryMajor"] = pitch_category(pitch_type)

        by_band: dict[str, list[dict]] = {}
        for e in type_entries:
            by_band.setdefault(e.get("band", ""), []).append(e)
        band_list = []
        for band, band_entries in by_band.items():
            band_stats = _agg_pitch_group(band_entries)
            band_stats["band"] = band
            band_list.append(band_stats)
        band_list.sort(key=lambda b: -(b["count"] or 0))
        stats["byVelocityBand"] = band_list

        result.append(stats)

    # 球種詳細の並びは「中カテゴリでまとめた順」に統一する（count順だとシーズンごと・
    # 選手ごとに順序がバラつき、プルダウン等の見た目が安定しないため）
    result.sort(key=lambda r: _pitch_sort_key(r["pitchType"]))
    return result


def build_pitch_type_breakdown(appearances: list[dict]) -> dict:
    """byPitchTypeの3系統（all/vsR/vsL）をまとめて返す"""
    return {
        "all": build_by_pitch_type(appearances),
        "vsR": build_by_pitch_type(appearances, hand_filter="R"),
        "vsL": build_by_pitch_type(appearances, hand_filter="L"),
    }


# ==================================================
# Section 2c. コースゾーン別集計（5x5=25分割、対左右投手の3系統: all/vsR/vsL）
# ==================================================
# 元データは run.py（NPB）・run_mlb.py（MLB）が日別JSONの各打者エントリに付与する
# "courseSplits"（試合×コースゾーン(行・列)×対戦投手利き腕、で事前集計済みのリスト）。
# ゾーンの行・列は投手側の座標ヒートマップ（pitcher-cards.htmlのhmBuildGridHtml/hmGetCell）
# と同じ5x5グリッド・同じ正規化式を使っており、x軸は打者自身の利き手で反転済み
# （+側=常に外角）。NPBは コース(Left)/コース(Top)、MLBはStatcastのplate_x/plate_zから、
# それぞれ同じ[-1,1]正規化空間に変換してからグリッド判定しているため、リーグ間でも
# ゾーンの意味は揃っている。

def build_by_course_zone(appearances: list[dict], hand_filter: str | None = None) -> list[dict]:
    """
    appearancesからbyCourseZone（all/vsR/vsLのうち1系統ぶん）を組み立てる。
    _agg_pitch_group()をそのまま流用する（courseSplitsのエントリはpitchSplitsと
    同じキー構成 n/pa/ab/h_/sw/ws/z/oz/zsw/ozsw を持つため）。
    """
    entries: list[dict] = []
    for ap in appearances:
        p = ap.get("player")
        if not isinstance(p, dict):
            continue
        for e in (p.get("courseSplits") or []):
            if hand_filter and e.get("h") != hand_filter:
                continue
            entries.append(e)

    by_cell: dict[tuple, list[dict]] = {}
    for e in entries:
        key = (e.get("row"), e.get("col"))
        by_cell.setdefault(key, []).append(e)

    result = []
    for (row, col), cell_entries in by_cell.items():
        if row is None or col is None:
            continue
        stats = _agg_pitch_group(cell_entries)
        stats["zoneRow"] = row
        stats["zoneCol"] = col
        result.append(stats)

    result.sort(key=lambda r: (r["zoneRow"], r["zoneCol"]))
    return result


def build_course_zone_breakdown(appearances: list[dict]) -> dict:
    """byCourseZoneの3系統（all/vsR/vsL）をまとめて返す"""
    return {
        "all": build_by_course_zone(appearances),
        "vsR": build_by_course_zone(appearances, hand_filter="R"),
        "vsL": build_by_course_zone(appearances, hand_filter="L"),
    }


# ==================================================
# Section 2d. カウント別集計（対左右投手の3系統: all/vsR/vsL）
# ==================================================
# 元データは run.py（NPB）・run_mlb.py（MLB）が日別JSONの各打者エントリに付与する
# "countSplits"（試合×カウント("{balls}-{strikes}")×対戦投手利き腕、で事前集計済みの
# リスト）。その球を最終球として受けた打席の結果を、そのカウントに紐付けている
# （打席途中の同じカウントでの見送り・ファウル等はスイング/空振り集計にのみ反映される）。

def build_by_count(appearances: list[dict], hand_filter: str | None = None) -> list[dict]:
    """appearancesからbyCount（all/vsR/vsLのうち1系統ぶん）を組み立てる"""
    entries: list[dict] = []
    for ap in appearances:
        p = ap.get("player")
        if not isinstance(p, dict):
            continue
        for e in (p.get("countSplits") or []):
            if hand_filter and e.get("h") != hand_filter:
                continue
            entries.append(e)

    by_count: dict[str, list[dict]] = {}
    for e in entries:
        by_count.setdefault(e.get("count", ""), []).append(e)

    result = []
    for count_key, count_entries in by_count.items():
        stats = _agg_pitch_group(count_entries)
        # 球数(count)は_agg_pitch_group側で正しく積み上げ済みなので、カウント表記
        # ("0-1"等)は別キー(countKey)に入れる（以前はここで"count"に上書きしてしまい、
        # 球数の値が失われていた）。
        stats["countKey"] = count_key
        result.append(stats)

    # ストライク数でまとめ、その中でボール数昇順に並べる（0-0,1-0,2-0,3-0 → 0-1,1-1,2-1,3-1 →
    # 0-2,1-2,2-2,3-2）。batter-cards.html側でストライク数ごとのグループヘッダーを出すのに使う。
    def _count_sort_key(r):
        try:
            b, s = r["countKey"].split("-")
            return (int(s), int(b))
        except (ValueError, AttributeError):
            return (99, 99)
    result.sort(key=_count_sort_key)
    return result


def build_count_breakdown(appearances: list[dict]) -> dict:
    """byCountの3系統（all/vsR/vsL）をまとめて返す"""
    return {
        "all": build_by_count(appearances),
        "vsR": build_by_count(appearances, hand_filter="R"),
        "vsL": build_by_count(appearances, hand_filter="L"),
    }


# ==================================================
# Section 2e. ランナー状況別集計（対左右投手の3系統: all/vsR/vsL）
# ==================================================
# 元データは run_mlb.py が日別JSONの各打者エントリに付与する"situationSplits"
# （試合×ランナー状況（走者なし/1塁/2塁/3塁/1・2塁/1・3塁/2・3塁/満塁）×対戦投手利き腕、
# で事前集計済みのリスト）。get_on_base_situation()による8状態の判定をそのまま使う。
# NPB側（run.py）は現状、生データにランナー状況の情報が含まれていないため未対応
# （situationSplitsを出力しない＝NPBではbySituationは空になる）。

def build_by_situation(appearances: list[dict], hand_filter: str | None = None) -> list[dict]:
    """appearancesからbySituation（all/vsR/vsLのうち1系統ぶん）を組み立てる"""
    entries: list[dict] = []
    for ap in appearances:
        p = ap.get("player")
        if not isinstance(p, dict):
            continue
        for e in (p.get("situationSplits") or []):
            if hand_filter and e.get("h") != hand_filter:
                continue
            entries.append(e)

    by_situation: dict[str, list[dict]] = {}
    for e in entries:
        by_situation.setdefault(e.get("situation", ""), []).append(e)

    result = []
    for situation_key, situation_entries in by_situation.items():
        stats = _agg_pitch_group(situation_entries)
        stats["situation"] = situation_key
        result.append(stats)

    # 表示順を「走者なし→1塁→2塁→3塁→1・2塁→1・3塁→2・3塁→満塁」に揃える
    _SITUATION_ORDER = ["走者なし", "走者1塁", "走者2塁", "走者3塁",
                         "走者1・2塁", "走者1・3塁", "走者2・3塁", "満塁"]
    result.sort(key=lambda r: _SITUATION_ORDER.index(r["situation"]) if r["situation"] in _SITUATION_ORDER else 99)
    return result


def build_situation_breakdown(appearances: list[dict]) -> dict:
    """bySituationの3系統（all/vsR/vsL）をまとめて返す"""
    return {
        "all": build_by_situation(appearances),
        "vsR": build_by_situation(appearances, hand_filter="R"),
        "vsL": build_by_situation(appearances, hand_filter="L"),
    }


# ==================================================
# Section 3. 試合ログ
# ==================================================

def build_game_log_rows_batter(appearances: list[dict]) -> list[dict]:
    rows = []
    for ap in appearances:
        p, g, side = ap["player"], ap["game"], ap["side"]
        opp = g.get("away") if side == "home" else g.get("home")
        ab = p.get("ab") or 0
        rows.append({
            "date": ap["date"], "opponent": opp, "opp_hand": ap.get("opp_pitcher_hand"),
            "order": p.get("order"), "pos": _clean_pos(p.get("pos")),
            "pa": p.get("pa"), "ab": p.get("ab"), "h": p.get("h"), "hr": p.get("hr"),
            "bb": p.get("bb"), "k": p.get("k"), "rbi": p.get("rbi"), "sb": p.get("sb"),
            "avg_game": round((p.get("h") or 0) / ab, 3) if ab > 0 else None,
            "obp": p.get("obp"), "slg": p.get("slg"), "ops": p.get("ops"),
            "abs": p.get("abs"),
        })
    return rows


# ==================================================
# Section 4b. シーズン成績ピボット用の拡張ランキング
# ==================================================
# 「シーズン成績」ピボット表（batter-cards.html）は、年度×対左右×球種大/中/詳細×球速帯を
# 自由に組み合わせて内訳を表示できるため、組み合わせごとに別々の母集団で順位を出す必要がある。
# ここでは4種類の母集団を用意する:
#   1) 対左右のみ（球種の絞り込みなし）→ splitRankings（overall/vsR/vsL）
#   2) 球種詳細単位            → byPitchTypeの各エントリの"rankings"（対左右population別）
#   3) 球種中/大カテゴリ単位    → byPitchCategoryMid/Majorの各エントリの"rankings"
#   4) 球種詳細×球速帯単位     → byVelocityBandの各エントリの"rankings"
# どの母集団も資格打席は統一してPIVOT_RANK_MIN_PA(=15)を使う（シーズン全体のKPIグリッド用
# ランキング＝RANK_MIN_PA(=100)とは別物。細かい内訳ほど閾値を緩める必要があるため）。
PIVOT_RANK_MIN_PA = 15

# (英語キー, 高いほど良いか)。打者にとって「良い」方向で統一する
# （K%・空振り率・chase%は低いほど良いので higher_is_better=False）。
_RANK_SPECS_EN = [
    ("avg", True), ("obp", True), ("slg", True), ("ops", True), ("hr", True), ("rbi", True),
    ("k_pct", False), ("bb_pct", True),
    ("whiff_pct", False), ("chase_pct", False), ("contact_pct", True),
]


def _compute_group_rankings(pools: dict[str, list[tuple[str, dict]]],
                             rank_min_pa: float = PIVOT_RANK_MIN_PA) -> dict[str, dict]:
    """任意のグループ分け（球種名・カテゴリ名・"球種|球速帯"キーなど）について、
    グループごとに独立した母集団で_RANK_SPECS_ENの各指標を順位付けする共通ロジック。
    pools: {グループキー: [(選手名, 統計オブジェクト(英語キー)), ...]}
    戻り値: {グループキー: {選手名: {metric: {"rank","total"}}}}
    """
    result: dict[str, dict] = {}
    for group_key, entries in pools.items():
        qualified = [(name, stat) for name, stat in entries if (stat.get("pa") or 0) >= rank_min_pa]
        group_result: dict[str, dict] = {}
        for metric, higher_is_better in _RANK_SPECS_EN:
            valid = [(name, stat[metric]) for name, stat in qualified if stat.get(metric) is not None]
            valid.sort(key=lambda x: -x[1] if higher_is_better else x[1])
            total = len(valid)
            for rank, (name, _) in enumerate(valid, start=1):
                group_result.setdefault(name, {})[metric] = {"rank": rank, "total": total}
        result[group_key] = group_result
    return result


def compute_split_rankings(stat_by_player: dict[str, dict], rank_min_pa: float = PIVOT_RANK_MIN_PA) -> dict:
    """対左右のみ（球種の絞り込み無し）の1系統ぶん（overall/vsR/vsLのいずれか）について、
    全打者を1つの母集団として順位付けする。
    stat_by_player: {選手名: 統計オブジェクト(英語キー、season.overallやsplits[hand]と同じ形)}
    戻り値: {選手名: {metric: {"rank","total"}}}
    """
    pools = {"_": list(stat_by_player.items())}
    return _compute_group_rankings(pools, rank_min_pa).get("_", {})


def compute_pitch_type_rankings(pt_lists_by_player: dict[str, list[dict]],
                                 rank_min_pa: float = PIVOT_RANK_MIN_PA) -> dict:
    """
    byPitchTypeの1系統（all/vsR/vsLのいずれか）ぶんについて、球種詳細ごとに
    打率・出塁率・長打率・OPS・本塁打・打点・K%・BB%・空振り率・chase%・contact%の順位を計算する。
    pt_lists_by_player: {選手名: build_by_pitch_type()の戻り値（1系統ぶん）}
    戻り値: {選手名: {球種名: {metric: {"rank","total"}}}}
    """
    pools: dict[str, list[tuple[str, dict]]] = {}
    for name, pt_list in pt_lists_by_player.items():
        for pt in pt_list:
            pools.setdefault(pt.get("pitchType", ""), []).append((name, pt))
    by_group = _compute_group_rankings(pools, rank_min_pa)
    result: dict[str, dict] = {name: {} for name in pt_lists_by_player}
    for pitch_type, by_name in by_group.items():
        for name, rk in by_name.items():
            result[name][pitch_type] = rk
    return result


def compute_pitch_band_rankings(pt_lists_by_player: dict[str, list[dict]],
                                 rank_min_pa: float = PIVOT_RANK_MIN_PA) -> dict:
    """
    球種詳細×球速帯（例:「ストレート」×「速い」）単位で順位を計算する。球速帯の閾値は
    球種コードごとに異なる（velo_band_label参照）ため、球種をまたいで同じ「速い」を
    比較するのではなく、必ず同じ球種内で比較する。
    pt_lists_by_player: {選手名: build_by_pitch_type()の戻り値（1系統ぶん）}
    戻り値: {選手名: {球種名: {球速帯: {metric: {"rank","total"}}}}}
    """
    pools: dict[tuple[str, str], list[tuple[str, dict]]] = {}
    for name, pt_list in pt_lists_by_player.items():
        for pt in pt_list:
            pitch_type = pt.get("pitchType", "")
            for band in (pt.get("byVelocityBand") or []):
                key = (pitch_type, band.get("band", ""))
                pools.setdefault(key, []).append((name, band))
    by_group = _compute_group_rankings(pools, rank_min_pa)
    result: dict[str, dict] = {name: {} for name in pt_lists_by_player}
    for (pitch_type, band_label), by_name in by_group.items():
        for name, rk in by_name.items():
            result[name].setdefault(pitch_type, {})[band_label] = rk
    return result


def compute_category_rankings(cat_map_by_player: dict[str, dict[str, dict]],
                               rank_min_pa: float = PIVOT_RANK_MIN_PA) -> dict:
    """
    球種中カテゴリ or 球種大カテゴリ単位で順位を計算する（compute_pitch_type_rankingsの
    カテゴリ版）。
    cat_map_by_player: {選手名: {カテゴリ名: 集約統計オブジェクト}}（build_by_pitch_categoryの戻り値）
    戻り値: {選手名: {カテゴリ名: {metric: {"rank","total"}}}}
    """
    pools: dict[str, list[tuple[str, dict]]] = {}
    for name, cat_map in cat_map_by_player.items():
        for cat_name, stat in cat_map.items():
            pools.setdefault(cat_name, []).append((name, stat))
    by_group = _compute_group_rankings(pools, rank_min_pa)
    result: dict[str, dict] = {name: {} for name in cat_map_by_player}
    for cat_name, by_name in by_group.items():
        for name, rk in by_name.items():
            result[name][cat_name] = rk
    return result


def _aggregate_pt_group_py(entries: list[dict]) -> dict:
    """既に集約済みのbyPitchType detail単位のエントリのリストを、さらに1つの統計オブジェクトに
    集約する（球種中/大カテゴリでのグルーピング用）。batter-cards.html側のJS版
    aggregatePitchEntries()と同じ考え方（打率・本塁打は安打/打数/本塁打の積み上げから正確に、
    長打率・出塁率・OPS・K%・BB%・空振り率・chase%・contact%は打席数で加重平均した近似値）。"""
    if not entries:
        return {}
    pa = sum(e.get("pa") or 0 for e in entries)
    ab = sum(e.get("ab") or 0 for e in entries)
    h = sum(e.get("h") or 0 for e in entries)
    hr = sum(e.get("hr") or 0 for e in entries)
    has_rbi = any(e.get("rbi") is not None for e in entries)
    rbi = sum(e.get("rbi") or 0 for e in entries) if has_rbi else None

    def _wavg(key, digits):
        num = sum((e.get(key) or 0) * (e.get("pa") or 0) for e in entries if e.get(key) is not None)
        den = sum((e.get("pa") or 0) for e in entries if e.get(key) is not None)
        return round(num / den, digits) if den > 0 else None

    return {
        "pa": pa, "ab": ab, "h": h, "hr": hr, "rbi": rbi,
        "avg": round(h / ab, 3) if ab > 0 else None,
        "slg": _wavg("slg", 3), "obp": _wavg("obp", 3), "ops": _wavg("ops", 3),
        "k_pct": _wavg("k_pct", 1), "bb_pct": _wavg("bb_pct", 1),
        "whiff_pct": _wavg("whiff_pct", 1), "chase_pct": _wavg("chase_pct", 1), "contact_pct": _wavg("contact_pct", 1),
    }


def build_by_pitch_category(pt_list: list[dict], category_key: str) -> dict[str, dict]:
    """byPitchType（1系統ぶん、detail単位のリスト）から、指定したカテゴリキー
    （"pitchCategoryMid" or "pitchCategoryMajor"）でグルーピングした集約統計を返す。
    戻り値: {カテゴリ名: 集約統計オブジェクト}"""
    groups: dict[str, list[dict]] = {}
    for pt in pt_list:
        key = pt.get(category_key) or "その他"
        groups.setdefault(key, []).append(pt)
    return {name: _aggregate_pt_group_py(entries) for name, entries in groups.items()}


# ==================================================
# Section 4. 順位算出・カテゴリタグ
# ==================================================

RANK_MIN_PA = 100  # 順位算出の資格打席（この打席数未満の選手は順位母集団から除外）

# 主定位置がこれに該当する選手は「投手」とみなし、打者一覧から除外する（NPBのみ該当。
# MLBは build_batter() が守備位置を常に空文字にしているため、この判定に引っかからない
# ＝現状MLB側では投手除外は機能しない。UDH制のため実害は小さいはずだが、位置データが
# 追加されたタイミングで見直すこと）。
PITCHER_POS_MARKERS = {"投", "(投)"}

# 打者名が数字だけ（例: "657675"）の行は、MLB側の選手名解決（playerid_reverse_lookup）
# が失敗し選手IDそのものにフォールバックしたケース。run_mlb.py側で修正済みだが、
# 今後も同様のフォールバックが起き得るため、ダッシュボードには出さないよう保険で除外する。
_NUMERIC_NAME_PAT = re.compile(r"^\d+$")

# (シーズン集計側の列名, numeric json側のキー名, 高いほど良いか)
_RANK_SPECS = [
    ("打率", "avg", True), ("出塁率", "obp", True), ("長打率", "slg", True), ("OPS", "ops", True),
    ("本塁打", "hr", True), ("盗塁", "sb", True), ("K%", "k_pct", False), ("BB%", "bb_pct", True),
]


def compute_batter_rankings(season_rows: list[dict], rank_min_pa: float = RANK_MIN_PA) -> dict:
    """打率・出塁率・長打率・OPS・本塁打・盗塁・K%・BB%の順位を算出する（KPIグリッド用。
    資格打席100のまま据え置き。ピボット表側は別途PIVOT_RANK_MIN_PA=15で計算する）。
    投手版と異なり役割による母集団分けは行わず、資格打席（rank_min_pa）以上の全打者を
    1つの母集団として順位付けする。
    戻り値: {選手名: {"avg":{"rank":n,"total":m}, ...}}
    """
    result = {row["選手名"]: {} for row in season_rows}
    pool = [row for row in season_rows if (row.get("打席") or 0) >= rank_min_pa]
    for jp_key, out_key, higher_is_better in _RANK_SPECS:
        valid = [(row["選手名"], row[jp_key]) for row in pool if row.get(jp_key) is not None]
        valid.sort(key=lambda x: -x[1] if higher_is_better else x[1])
        total = len(valid)
        for rank, (name, _) in enumerate(valid, start=1):
            result[name][out_key] = {"rank": rank, "total": total}
    return result


def compute_defense_rankings(oaa_by_player: dict[str, list[dict]]) -> dict:
    """
    守備OAA（outs_above_average）を、同じポジションでプレーした選手同士だけで
    比較して順位を付ける（ポジションによって守備機会・難易度が全く違うため、
    全選手一律で比較するのは意味がない。ポジションごとに独立した母集団で順位付けする）。

    oaa_by_player: {選手名: card["defense"]["oaa"]のリスト（各要素は{"pos":..,"outs_above_average":..,...}）}
    戻り値: {選手名: {ポジション名: {"rank": n, "total": m}}}
    """
    pools: dict[str, list[tuple[str, float]]] = {}
    for name, oaa_list in oaa_by_player.items():
        for entry in (oaa_list or []):
            pos = entry.get("pos")
            val = entry.get("outs_above_average")
            if pos is None or val is None:
                continue
            pools.setdefault(pos, []).append((name, val))

    result: dict[str, dict] = {name: {} for name in oaa_by_player}
    for pos, entries in pools.items():
        # outs_above_averageは高いほど良い（他球種別ランキングと同じ「高いほど良い」に統一）
        valid = sorted(entries, key=lambda x: -x[1])
        total = len(valid)
        for rank, (name, _) in enumerate(valid, start=1):
            result[name][pos] = {"rank": rank, "total": total}
    return result


def classify_batter_categories(card: dict) -> list:
    """通算成績から、選手一覧で絞り込みに使えるカテゴリタグを機械的に組み立てる"""
    cats = []
    if (card.get("hr") or 0) >= 20:
        cats.append("パワーヒッター")
    if (card.get("sb") or 0) >= 15:
        cats.append("走力型")
    bbpct, kpct = card.get("bb_pct_season"), card.get("k_pct_season")
    if bbpct is not None and bbpct >= 12:
        cats.append("選球眼型")
    if kpct is not None and kpct <= 15:
        cats.append("コンタクト型")
    return cats


def _slugify_name(name: str) -> str:
    """選手名からファイル名用のIDを作る（英数字以外はアンダースコアに置換）"""
    s = re.sub(r"[^\w]+", "_", name.strip().lower())
    return s.strip("_") or "unknown"


def load_mlb_defense_cache(path: str) -> dict:
    """
    run_mlb.py の --steps defense が出力する中間キャッシュxlsx（"OAA"/"SprintSpeed"/
    "CatcherPoptime"シート）を読み、正規化した選手名をキーに守備・走塁指標を返す。
    戻り値: {正規化した選手名: {"oaa":[...], "sprint_speed":float|None, "poptime":{...}|None}}

    - OAAシートは選手が複数ポジションでプレーしていると複数行に分かれるため、リストで持つ
      （合計値を1つに潰すと「どのポジションでの数値か」が失われるため）。
    - poptimeは捕手のみ該当（それ以外の選手は該当行が無いのでNoneのまま）。
    - 読み込みに失敗した場合は空dictを返す（呼び出し側は defense が従来通りnullになるだけで、
      パイプライン全体は止めない）。
    """
    result: dict = {}
    if not path or not os.path.isfile(path):
        return result
    try:
        wb = pd.ExcelFile(path)
    except Exception as e:
        print(f"  [WARN] 守備・走塁キャッシュの読み込みに失敗しました（{path}）: {e}")
        return result

    def _entry(name):
        key = _normalize_name(name).lower()
        return result.setdefault(key, {
            "oaa": [], "sprint_speed": None, "framing": None, "poptime": None,
        })

    if "OAA" in wb.sheet_names:
        for _, r in wb.parse("OAA").iterrows():
            name = r.get("name")
            if not name or (isinstance(name, float) and pd.isna(name)):
                continue
            _entry(name)["oaa"].append({
                "pos": r.get("pos"),
                "outs_above_average": None if pd.isna(r.get("outs_above_average")) else int(r.get("outs_above_average")),
                "fielding_runs_prevented": None if pd.isna(r.get("fielding_runs_prevented")) else int(r.get("fielding_runs_prevented")),
                "actual_success_rate": None if pd.isna(r.get("actual_success_rate")) else float(r.get("actual_success_rate")),
            })

    if "SprintSpeed" in wb.sheet_names:
        for _, r in wb.parse("SprintSpeed").iterrows():
            name = r.get("name")
            if not name or (isinstance(name, float) and pd.isna(name)):
                continue
            v = r.get("sprint_speed")
            _entry(name)["sprint_speed"] = None if pd.isna(v) else float(v)

    # run_mlb.py の fetch_catcher_framing() / fetch_catcher_poptime() は、Statcastの生の列名
    # ではなく簡略化した列名（framing_runs/framing_pct、arm_strength/exchange_time/pop_2b/pop_3b等）
    # で既にxlsxに書き出している。ここではその簡略化後の列名をそのまま読む。
    if "CatcherFraming" in wb.sheet_names:
        for _, r in wb.parse("CatcherFraming").iterrows():
            name = r.get("name")
            if not name or (isinstance(name, float) and pd.isna(name)):
                continue
            _entry(name)["framing"] = {
                "pitches": None if pd.isna(r.get("pitches")) else int(r.get("pitches")),
                "framing_runs": None if pd.isna(r.get("framing_runs")) else float(r.get("framing_runs")),
                "framing_pct": None if pd.isna(r.get("framing_pct")) else float(r.get("framing_pct")),
            }

    if "CatcherPoptime" in wb.sheet_names:
        for _, r in wb.parse("CatcherPoptime").iterrows():
            name = r.get("name")
            if not name or (isinstance(name, float) and pd.isna(name)):
                continue
            _entry(name)["poptime"] = {
                "arm_strength": None if pd.isna(r.get("arm_strength")) else float(r.get("arm_strength")),
                "exchange_time": None if pd.isna(r.get("exchange_time")) else float(r.get("exchange_time")),
                "pop_2b": None if pd.isna(r.get("pop_2b")) else float(r.get("pop_2b")),
                "pop_3b": None if pd.isna(r.get("pop_3b")) else float(r.get("pop_3b")),
            }

    return result


def determine_season_year(all_data: dict) -> str:
    """games/json内の全日付のうち最も多い年を、このシーズンの年とみなす
    （通常は単一年のデータしか渡らないが、年またぎのデータが混じっていても
    多数派の年を採用することで壊れないようにしておく）"""
    dates = [d for d in all_data.keys() if d != "highlights" and not str(d).startswith("_")]
    years = [str(d)[:4] for d in dates if len(str(d)) >= 4 and str(d)[:4].isdigit()]
    if not years:
        return str(datetime.date.today().year)
    return Counter(years).most_common(1)[0][0]


def _to_stat_obj(s: dict) -> dict:
    """calc_season_batter_stats()の日本語キー辞書を、numeric_json用の英語キー
    「成績オブジェクト」（batter_card_schema.md参照）に変換する"""
    return {
        "games": s.get("試合数"), "pa": s.get("打席"), "ab": s.get("打数"),
        "h": s.get("安打"), "hr": s.get("本塁打"), "bb": s.get("四死球"),
        "k": s.get("三振"), "rbi": s.get("打点"), "sb": s.get("盗塁"), "cs": s.get("盗塁死"),
        "avg": s.get("打率"), "obp": s.get("出塁率"), "slg": s.get("長打率"), "ops": s.get("OPS"),
        "k_pct": s.get("K%"), "bb_pct": s.get("BB%"),
        "chase_pct": s.get("O-Swing%"), "contact_pct": s.get("Z-Swing%"), "whiff_pct": s.get("whiff%"),
    }


# ==================================================
# Section 5. メイン: xlsx / numeric json 出力
# ==================================================

def export_llm_input_batter_xlsx(games_json_dir: str, out_path: str, min_pa: float = 0.0,
                                  target_names: list[str] | None = None,
                                  numeric_json_dir: str | None = None,
                                  mlb_defense_xlsx: str | None = None) -> str:
    """
    numeric_json_dir を指定すると、xlsxに加えて選手ごとの数値データJSON
    （batter_cards_numeric/{選手ID}.json）と選手一覧 index.json も書き出す。
    これらはbatter-cards.html側がそのまま読み込む「数値だけ」のデータ。

    mlb_defense_xlsx: run_mlb.py の --steps defense が出力する守備OAA・
    スプリントスピードの中間キャッシュxlsx（"OAA"/"SprintSpeed"シート）。
    指定すると、名前が一致する選手の numeric_json の defense.oaa / defense.sprint_speed
    を埋める（MLBのみ。NPBでは通常この引数を渡さないのでNoneのまま＝従来通り）。
    """
    defense_cache = load_mlb_defense_cache(mlb_defense_xlsx) if mlb_defense_xlsx else {}

    all_data = load_daily_games(games_json_dir)
    names = set(target_names) if target_names else build_all_batter_names(all_data)
    season_year = determine_season_year(all_data)

    season_rows, split_rows, gamelog_rows = [], [], []
    numeric_cards: dict[str, dict] = {}

    for name in sorted(names):
        try:
            if _NUMERIC_NAME_PAT.match(name.strip()):
                print(f"  [SKIP] {name}: 選手名が数字のみ（名前解決失敗の疑い）のため除外")
                continue

            appearances = build_appearances_batter(all_data, name)
            if not appearances:
                continue

            season = calc_season_batter_stats(appearances)
            season["選手名"] = name

            if (season.get("打席") or 0) < min_pa:
                continue

            pos_raw = determine_primary_position(appearances)
            if pos_raw in PITCHER_POS_MARKERS:
                continue
            pos = _clean_pos(pos_raw)  # 表示・フィルタ用にカッコを剥がした値（"(三)"→"三"）

            vs_r = calc_season_batter_stats(appearances, hand_filter="R")
            vs_l = calc_season_batter_stats(appearances, hand_filter="L")

            season["主定位置"] = pos

            # 直近の出場からチーム名を推定（home/awayどちら側だったかで判定）
            last_ap = appearances[-1]
            last_game = last_ap.get("game") or {}
            team = last_game.get(last_ap.get("side")) if isinstance(last_game, dict) else None
            season["所属チーム"] = team

            season_rows.append(season)
            split_rows.append({"選手名": name, "対戦": "対右投手", **vs_r})
            split_rows.append({"選手名": name, "対戦": "対左投手", **vs_l})

            game_log_dicts = build_game_log_rows_batter(appearances)
            for g in game_log_dicts:
                gamelog_rows.append({"選手名": name, **g})

            if numeric_json_dir:
                cs = season.get("盗塁死")  # NoneならNPB等cs未対応、0以上ならMLBで実際に集計された値
                sb_n = season.get("盗塁") or 0
                sb_success_pct = (
                    round(sb_n / (sb_n + cs) * 100, 1) if cs is not None and (sb_n + cs) > 0 else None
                )
                defense_entry = defense_cache.get(_normalize_name(name).lower(), {})
                pt_breakdown = build_pitch_type_breakdown(appearances)
                numeric_cards[name] = {
                    "team": team, "pos": pos,
                    "overall": _to_stat_obj(season),
                    "splits": {"vsR": _to_stat_obj(vs_r), "vsL": _to_stat_obj(vs_l)},
                    "byPitchType": pt_breakdown,
                    "byPitchCategoryMid": {
                        pop: build_by_pitch_category(pt_breakdown.get(pop) or [], "pitchCategoryMid")
                        for pop in ("all", "vsR", "vsL")
                    },
                    "byPitchCategoryMajor": {
                        pop: build_by_pitch_category(pt_breakdown.get(pop) or [], "pitchCategoryMajor")
                        for pop in ("all", "vsR", "vsL")
                    },
                    "byCourseZone": build_course_zone_breakdown(appearances),
                    "byCount": build_count_breakdown(appearances),
                    "bySituation": build_situation_breakdown(appearances),
                    "game_log": game_log_dicts,
                    "defense": {
                        # 守備率・レンジファクターは、現状のrun.py/run_mlb.pyに守備成績の
                        # スクレイピング/取得が実装されていないため取得不可（将来対応まではNone）。
                        "fielding_pct": None,
                        "range_factor": None,
                        # 盗塁死・盗塁成功率はMLBのみ算出可能（run_mlb.pyが試合全体から走者を
                        # 正しく特定して集計する。NPBの打者エントリには"cs"が無いためNoneのまま）。
                        "cs": cs,
                        "sb_success_pct": sb_success_pct,
                        # 以下はMLBのみ、run_mlb.py --steps defense のキャッシュ(mlb_defense_xlsx)が
                        # 渡された場合にだけ埋まる。無ければ全てNone/空リストのまま。
                        "oaa": defense_entry.get("oaa", []),
                        "sprint_speed": defense_entry.get("sprint_speed"),
                        "catcher_framing": defense_entry.get("framing"),
                        "catcher_poptime": defense_entry.get("poptime"),
                    },
                    # rankingsはこの後、全選手分揃ってから付与する
                }
        except Exception as e:
            print(f"  [SKIP] {name}: 集計中にエラーのためスキップ({type(e).__name__}: {e})")
            continue

    rankings = compute_batter_rankings(season_rows)
    for row in season_rows:
        rk = rankings.get(row["選手名"], {})
        for jp_key, out_key, _ in _RANK_SPECS:
            row[f"{jp_key}_順位"] = rk.get(out_key, {}).get("rank")
            row[f"{jp_key}_順位_母数"] = rk.get(out_key, {}).get("total")

    # ── シーズン成績ピボット用の拡張ランキング（PIVOT_RANK_MIN_PA=15を資格打席として使う） ──
    # 1) 対左右のみ（球種の絞り込み無し）: overall/vsR/vsLそれぞれ独立した母集団で計算し、
    #    season["splitRankings"]に付与する。
    for population, stat_key in (("all", "overall"), ("vsR", "vsR"), ("vsL", "vsL")):
        stat_by_player = {
            name: (card["overall"] if stat_key == "overall" else card["splits"][stat_key])
            for name, card in numeric_cards.items()
        }
        split_rk = compute_split_rankings(stat_by_player)
        for name, card in numeric_cards.items():
            card.setdefault("splitRankings", {})[population] = split_rk.get(name, {})

    # 2) 球種詳細ランキング（all/vsR/vsLそれぞれ独立した母集団で計算し、対応するbyPitchType
    #    の各球種オブジェクトに rankings として付与する）
    for population in ("all", "vsR", "vsL"):
        pt_lists_by_player = {
            name: (card["byPitchType"].get(population) or [])
            for name, card in numeric_cards.items()
        }
        pt_rankings = compute_pitch_type_rankings(pt_lists_by_player)
        for name, card in numeric_cards.items():
            for pt in (card["byPitchType"].get(population) or []):
                pt["rankings"] = pt_rankings.get(name, {}).get(pt.get("pitchType", ""), {})

        # 3) 球種詳細×球速帯ランキング（例:「ストレート」×「速い」）。byVelocityBandの
        #    各エントリにrankingsを付与する（球種ごとに独立した母集団で比較するため、
        #    球種をまたいだ「速い」同士の比較にはならない）。
        band_rankings = compute_pitch_band_rankings(pt_lists_by_player)
        for name, card in numeric_cards.items():
            for pt in (card["byPitchType"].get(population) or []):
                pitch_type = pt.get("pitchType", "")
                for band in (pt.get("byVelocityBand") or []):
                    band["rankings"] = band_rankings.get(name, {}).get(pitch_type, {}).get(band.get("band", ""), {})

        # 4) 球種中/大カテゴリランキング（byPitchCategoryMid/Majorの各エントリにrankingsを付与）
        for cat_field, cat_key in (("byPitchCategoryMid", "pitchCategoryMid"), ("byPitchCategoryMajor", "pitchCategoryMajor")):
            cat_map_by_player = {
                name: (card[cat_field].get(population) or {})
                for name, card in numeric_cards.items()
            }
            cat_rankings = compute_category_rankings(cat_map_by_player)
            for name, card in numeric_cards.items():
                for cat_name, stat in (card[cat_field].get(population) or {}).items():
                    stat["rankings"] = cat_rankings.get(name, {}).get(cat_name, {})

    # 守備OAAのポジション別ランキング（KPI表示用。同じポジションの選手同士だけで比較する）
    oaa_by_player = {name: card["defense"].get("oaa", []) for name, card in numeric_cards.items()}
    defense_rankings = compute_defense_rankings(oaa_by_player)
    for name, card in numeric_cards.items():
        for entry in card["defense"].get("oaa", []):
            entry["rank"] = defense_rankings.get(name, {}).get(entry.get("pos"), {})

    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame(season_rows).to_excel(writer, sheet_name="シーズン集計", index=False)
        pd.DataFrame(split_rows).to_excel(writer, sheet_name="対左右投手別", index=False)
        pd.DataFrame(gamelog_rows).to_excel(writer, sheet_name="試合ログ", index=False)

    if numeric_json_dir:
        os.makedirs(numeric_json_dir, exist_ok=True)
        index_players = []
        for name, season_obj in numeric_cards.items():
            season_obj["rankings"] = rankings.get(name, {})
            player_id = _slugify_name(name)
            path = os.path.join(numeric_json_dir, f"{player_id}.json")

            # 既存ファイルがあれば読み込み、seasons辞書の該当年だけ更新する
            # （他の年のデータ・MLBの過去シーズン分などを上書きしないため）
            existing: dict = {}
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        existing = json.load(f)
                except Exception as e:
                    print(f"  [WARN] {player_id}.json の既存データ読み込みに失敗（新規として扱います）: {e}")
                    existing = {}

            seasons = existing.get("seasons") if isinstance(existing.get("seasons"), dict) else {}
            seasons[season_year] = season_obj
            years_sorted = sorted(seasons.keys(), reverse=True)
            latest_year = years_sorted[0]
            latest = seasons[latest_year]
            latest_overall = latest.get("overall", {})

            categories = classify_batter_categories({
                "hr": latest_overall.get("hr"), "sb": latest_overall.get("sb"),
                "bb_pct_season": latest_overall.get("bb_pct"),
                "k_pct_season": latest_overall.get("k_pct"),
            })

            full_card = {
                "name": name,
                "team": latest.get("team"), "pos": latest.get("pos"),
                "years": years_sorted, "latestYear": latest_year,
                # ── 後方互換ミラー：最新シーズンのoverallと同一内容 ──
                "games": latest_overall.get("games"), "pa": latest_overall.get("pa"),
                "ab": latest_overall.get("ab"), "h": latest_overall.get("h"),
                "hr": latest_overall.get("hr"), "bb": latest_overall.get("bb"),
                "k": latest_overall.get("k"), "rbi": latest_overall.get("rbi"),
                "sb": latest_overall.get("sb"),
                "avg": latest_overall.get("avg"), "obp": latest_overall.get("obp"),
                "slg": latest_overall.get("slg"), "ops": latest_overall.get("ops"),
                "k_pct_season": latest_overall.get("k_pct"),
                "bb_pct_season": latest_overall.get("bb_pct"),
                "chase_pct_season": latest_overall.get("chase_pct"),
                "contact_pct_season": latest_overall.get("contact_pct"),
                "whiff_pct_season": latest_overall.get("whiff_pct"),
                "rankings": latest.get("rankings", {}),
                "game_log": latest.get("game_log", []),
                "categories": categories,
                "seasons": seasons,
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(full_card, f, ensure_ascii=False)

            index_players.append({
                "id": player_id, "name": name, "team": full_card.get("team"), "pos": full_card.get("pos"),
                "categories": categories,
                "games": full_card.get("games"), "pa": full_card.get("pa"),
                "avg": full_card.get("avg"), "obp": full_card.get("obp"), "slg": full_card.get("slg"),
                "ops": full_card.get("ops"),
                "hr": full_card.get("hr"), "rbi": full_card.get("rbi"), "sb": full_card.get("sb"),
                "k_pct": full_card.get("k_pct_season"), "bb_pct": full_card.get("bb_pct_season"),
            })
        with open(os.path.join(numeric_json_dir, "index.json"), "w", encoding="utf-8") as f:
            json.dump({"players": index_players}, f, ensure_ascii=False, indent=2)
        print(f"  数値JSON: {len(index_players)}選手分を {numeric_json_dir} に出力（対象シーズン: {season_year}）")

    return out_path


# ==================================================
# Section 6. CLI
# ==================================================

def parse_args():
    p = argparse.ArgumentParser(description="LLM入力用xlsx（シーズン打者データ）を生成する")
    p.add_argument("--games-json-dir", required=True, help="games/json/{date}.json が入っているディレクトリ")
    p.add_argument("--out", required=True, help="出力xlsxのパス")
    p.add_argument("--min-pa", type=float, default=0.0, help="この打席数以上の打者のみ対象にする（デフォルト0=全員）")
    p.add_argument("--players", nargs="*", default=None, help="対象選手名を絞り込む場合はスペース区切りで指定（省略時は全打者）")
    p.add_argument("--numeric-json-dir", default=None,
                    help="指定すると、選手ごとの数値データJSON（batter_cards_numeric/配下）とindex.jsonも出力する")
    p.add_argument("--mlb-defense-xlsx", default=None,
                    help="run_mlb.pyの--steps defenseが出力する守備OAA・走塁スプリントスピードの中間キャッシュxlsxのパス（MLBのみ）")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    out = export_llm_input_batter_xlsx(
        games_json_dir=args.games_json_dir,
        out_path=args.out,
        min_pa=args.min_pa,
        target_names=args.players,
        numeric_json_dir=args.numeric_json_dir,
        mlb_defense_xlsx=args.mlb_defense_xlsx,
    )
    print(f"✅ 出力完了: {out}")
