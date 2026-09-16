# -*- coding: utf-8 -*-
"""
export_llm_input.py
====================
index.html 側でクライアントサイド（JavaScript）に実装されている、投手の
シーズン集計・球種別詳細・コース分布・カウント別パターンの集計ロジックを
Pythonに移植し、Claudeへ渡すLLM入力用xlsx（縦持ち・複数シート）を生成する。

ポート元:
  - calcSeasonStats()      → シーズン投手成績（防御率・K-BB%など）
  - _aggregateSeasonMix()  → 球種別シーズン集計（cbsのマージを含む）
  - 9分割コース分布ロジック → コース分布（対右/対左、ゾーン内9マス）

入力:
  run.py の run_dashboard() が既に生成している「日別ダッシュボードJSON」
  （games/json/{date}.json）を、対象期間ぶんすべて読み込んで使う。
  players/json/{選手名}.json（MLB専用のキャッシュ）は使わない
  → NPBはこの日別JSONの積み上げだけでシーズン集計が可能なため。

出力:
  1つのxlsxに以下のシートを縦持ち（long形式）で書き出す。
    - シーズン集計       : 1行 = 1投手（選手名・投球成績に加え「所属チーム」「役割」列を含む）
    - 球種別詳細         : 1行 = 1投手 × 1球種。全体/対右/対左それぞれの空振り率・
                          ゾーン外スイング率・ストライク率・ゾーン率・GB%(ゴロ率)に加え、
                          指標ごとの{指標}_順位・{指標}_順位_母数を含む（役割別・
                          全体は投球回条件、対右/対左はさらに対戦数条件を満たす投手の
                          中での順位。compute_pitch_rankings()参照）
    - コース分布         : 1行 = 1投手 × 対戦打者(右/左) × ゾーン(1-9、全球種合算)
    - コースグリッド25分割 : 1行 = 1投手 × 球種 × 対戦打者(右/左) × セル(5x5、ボール域込み)
    - カウント別パターン : 1行 = 1投手 × 球種 × カウント状況

使い方:
    python export_llm_input.py \
        --games-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/games/json" \
        --out "data/datamart/llm_input/プロ野球_2026_投手データ.xlsx" \
        --min-ip 10

    # 投球回フィルタなしで全投手を対象にする場合は --min-ip を省略（デフォルト0）
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import re
from collections import defaultdict

import pandas as pd


# ==================================================
# Section 1. 日別JSONの読み込み・appearances構築
#   index.html の _doOpenPlayerPage() 相当
# ==================================================

def load_daily_games(games_json_dir: str, verbose: bool = True) -> dict:
    """games/json/{date}.json を全部読み込み、{date: [game, ...]} にまとめる。
    壊れた/想定外の形式のファイルやゲームエントリはスキップし、警告を出して処理を継続する。
    """
    all_data: dict[str, list] = {}
    paths = sorted(glob.glob(os.path.join(games_json_dir, "*.json")))
    if not paths:
        raise FileNotFoundError(f"games json が見つかりません: {games_json_dir}")

    def _add_games(date_key: str, games) -> None:
        if not isinstance(games, list):
            if verbose:
                print(f"  [SKIP] {date_key}: gamesがlistではない型({type(games).__name__})のためスキップ")
            return
        valid_games = [g for g in games if isinstance(g, dict)]
        skipped = len(games) - len(valid_games)
        if skipped > 0 and verbose:
            print(f"  [SKIP] {date_key}: dict以外のgameエントリを{skipped}件スキップ")
        all_data.setdefault(date_key, []).extend(valid_games)

    for path in paths:
        date = os.path.splitext(os.path.basename(path))[0]
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            if verbose:
                print(f"  [SKIP] {os.path.basename(path)}: 読み込み失敗({e})")
            continue

        if not isinstance(payload, dict):
            if verbose:
                print(f"  [SKIP] {os.path.basename(path)}: 想定外のトップレベル型({type(payload).__name__})")
            continue

        if date in payload:
            _add_games(date, payload[date])
        else:
            for k, v in payload.items():
                if k == "highlights" or (isinstance(k, str) and k.startswith("_")):
                    continue
                _add_games(k, v)
    return all_data


def _normalize_pitcher_name(name: str) -> str:
    """
    表記ゆれ吸収: 'Nola, Aaron' 形式と 'Aaron Nola' 形式が元データに混在しているため、
    'Last, First' 形式を 'First Last' 形式に統一して同一選手として扱えるようにする。
    """
    if not isinstance(name, str):
        return name
    name = name.strip()
    if "," in name:
        parts = name.split(",", 1)
        last, first = parts[0].strip(), parts[1].strip()
        if last and first:
            return f"{first} {last}"
    return name


def build_all_pitcher_names(all_data: dict) -> set[str]:
    """全登場投手名を収集（表記ゆれを正規化・壊れたエントリはスキップ）"""
    names = set()
    for date, games in all_data.items():
        if date == "highlights" or date.startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict):
                continue
            pitchers = g.get("pitchers")
            if not isinstance(pitchers, dict):
                continue
            for side in ("home", "away"):
                for p in (pitchers.get(side) or []):
                    if isinstance(p, dict) and p.get("name"):
                        names.add(_normalize_pitcher_name(p["name"]))
    return names


def build_appearances(all_data: dict, player_name: str) -> list[dict]:
    """特定投手の全登板 = appearances（index.htmlのappearances配列と同じ形）。
    表記ゆれ（'Last, First' / 'First Last'）を正規化してから比較する。壊れたエントリはスキップ。
    """
    appearances = []
    dates = sorted(d for d in all_data.keys() if d != "highlights" and not d.startswith("_"))
    name_norm = _normalize_pitcher_name(player_name).lower()
    for date in dates:
        for g in all_data.get(date, []):
            if not isinstance(g, dict):
                continue
            pitchers = g.get("pitchers")
            if not isinstance(pitchers, dict):
                continue
            for side in ("home", "away"):
                for p in (pitchers.get(side) or []):
                    if not isinstance(p, dict):
                        continue
                    raw_name = p.get("name") or ""
                    if _normalize_pitcher_name(raw_name).lower() == name_norm:
                        appearances.append({"date": date, "game": g, "side": side, "player": p})
    return appearances


# ==================================================
# Section 2. シーズン集計（calcSeasonStats のポート）
# ==================================================

def _ip_to_outs(ip_str) -> int:
    """'6.2' 形式（6回2/3）→ アウト数（6*3+2=20）"""
    s = str(ip_str or "0")
    whole, _, frac = s.partition(".")
    whole_n = int(whole) if whole.strip().lstrip("-").isdigit() else 0
    frac_n = 1 if frac == "1" else 2 if frac == "2" else 0
    return whole_n * 3 + frac_n


def _outs_to_ip_str(outs: int) -> str:
    whole = outs // 3
    rem = outs % 3
    return f"{whole}.{rem}"


def calc_season_stats(appearances: list[dict]) -> dict:
    """calcSeasonStats(mode='all') のポート。防御率・K-BB%などシーズン集計を1行返す"""
    appearances = [ap for ap in appearances if isinstance(ap.get("player"), dict)]
    if not appearances:
        return {"選手名": None, "登板数": 0, "投球回": "0.0", "奪三振": 0, "与四球": 0,
                "被安打": 0, "自責点": 0, "防御率": None, "K%": None, "BB%": None, "K-BB%": None,
                "空振り率": None, "ゾーン外スイング率": None, "ストライク率": None, "ゾーン率": None, "ゴロ率": None}

    total_ip_raw = sum(_ip_to_outs(ap["player"].get("ip")) for ap in appearances)
    total_k  = sum(ap["player"].get("k", 0) or 0 for ap in appearances)
    total_bb = sum(ap["player"].get("bb", 0) or 0 for ap in appearances)
    total_h  = sum(ap["player"].get("h", 0) or 0 for ap in appearances)
    total_er = sum(ap["player"].get("er", 0) or 0 for ap in appearances)

    ip_num = total_ip_raw / 3
    era = round(total_er / ip_num * 9, 2) if ip_num > 0 else None

    total_tbf = sum(
        (ap["player"].get("tbf") or (ap["player"].get("k", 0) or 0)
         + (ap["player"].get("bb", 0) or 0) + (ap["player"].get("h", 0) or 0))
        for ap in appearances
    )
    kpct   = round(total_k / total_tbf * 100, 1) if total_tbf > 0 else None
    bbpct  = round(total_bb / total_tbf * 100, 1) if total_tbf > 0 else None
    kbbpct = round((total_k - total_bb) / total_tbf * 100, 1) if total_tbf > 0 else None

    tot_pitches = swstr_sum = strike_sum = zone_sum = 0.0
    oswing_sum, oswing_count = 0.0, 0.0
    gb_sum, gb_count = 0.0, 0.0
    for ap in appearances:
        p = ap["player"]
        n = p.get("pitches", 0) or 0
        if n > 0:
            tot_pitches += n
            if p.get("swstr")  is not None: swstr_sum  += p["swstr"]  * n
            if p.get("strike") is not None: strike_sum += p["strike"] * n
            if p.get("zone")   is not None: zone_sum   += p["zone"]   * n
        if p.get("oSwing") is not None:
            # ゾーン外スイング率の本来の分母は「ゾーン外投球数」（全投球数ではない）。
            # 「0」は「ゾーン外投球が0球だった」という正当な値なのでそのまま使う。
            # oz_n自体が無い（まだ再生成していない古い形式の）試合は、投球数で代用すると
            # 単位が違う重み（投球数 vs ゾーン外投球数）が混在して古い試合が不当に重く
            # 扱われてしまうため、その試合はこの指標の集計から除外する。
            ow = p.get("oz_n")
            if ow is not None:
                oswing_sum += p["oSwing"] * ow
                oswing_count += ow
        if p.get("gbpct") is not None:
            # ゴロ率は「試合ごとのGB数」「試合ごとの打球数(bip_known)」をそのまま合計して
            # 割る（分子・分母の生の実数を積み上げる）。
            # bip_knownが「0」（その試合その球種で打球が1つも無かった）場合は正当な値
            # なのでそのまま使う。bip_known自体が無い古い形式の試合は、投球数で代用すると
            # 単位が違う重みが混ざって古い試合が不当に重く扱われるため、除外する。
            w = p.get("bip_known")
            if w is not None:
                gn = p.get("gb_n")
                gb_sum += gn if gn is not None else (p["gbpct"] / 100) * w
                gb_count += w

    # MLB限定の高度指標（avgEV/hardHitPct/barrelPct/xwoba/avgSpin/extension/vaa）
    # NPBデータには存在しないため、appearance側にキーがある分だけ拾う（単純平均。登板ごとの重みは球数ではなくフラットに）
    _adv_fields = ["avgEV", "hardHitPct", "barrelPct", "xwoba", "avgSpin", "extension", "vaa"]
    adv_sums = {f: 0.0 for f in _adv_fields}
    adv_counts = {f: 0 for f in _adv_fields}
    for ap in appearances:
        p = ap["player"]
        for f in _adv_fields:
            v = p.get(f)
            if v is not None:
                adv_sums[f] += v
                adv_counts[f] += 1
    adv_out = {
        f"season_{f}": (round(adv_sums[f] / adv_counts[f], 3) if adv_counts[f] > 0 else None)
        for f in _adv_fields
    }

    return {
        "選手名":   appearances[0]["player"]["name"] if appearances else None,
        "登板数":   len(appearances),
        "投球回":   _outs_to_ip_str(total_ip_raw),
        "奪三振":   total_k,
        "与四球":   total_bb,
        "被安打":   total_h,
        "自責点":   total_er,
        "防御率":   era,
        "K%":      kpct,
        "BB%":     bbpct,
        "K-BB%":   kbbpct,
        "空振り率":  round(swstr_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゾーン外スイング率": round(oswing_sum / oswing_count, 1) if oswing_count > 0 else None,
        "ストライク率": round(strike_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゾーン率":  round(zone_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゴロ率":   round(gb_sum / gb_count * 100, 1) if gb_count > 0 else None,
        "平均被打球速度": adv_out["season_avgEV"],       # MLBのみ。NPBはNone
        "ハードヒット率": adv_out["season_hardHitPct"],   # MLBのみ
        "バレル率":      adv_out["season_barrelPct"],     # MLBのみ
        "xwOBA":        adv_out["season_xwoba"],          # MLBのみ
        "平均回転数":     adv_out["season_avgSpin"],       # MLBのみ
        "エクステンション": adv_out["season_extension"],   # MLBのみ
        "VAA":          adv_out["season_vaa"],             # MLBのみ（縦の入射角）
    }


def calc_season_stats_by_hand(appearances: list[dict]) -> dict:
    """
    対右打者・対左打者それぞれの「本当のシーズン実数」を、試合ごとのpitStatVsR/pitStatVsL
    （各試合で計算済みの正しい分母＝bip_known/oz_nを持つ）を積み上げて算出する。

    球種比較テーブルの「合計」行に使う値。球種別内訳（season_pitch_detail）から
    再度加重平均するのではなく、この関数の値をそのまま使うことで、「一部の打席で球種が
    特定できず球種別テーブルには反映されない」といったケースでも実際のシーズン数値と
    ズレないようにする。

    戻り値: {"vsR": {"空振り率":..., "ゾーン外スイング率":..., "ストライク率":..., "ゾーン率":..., "ゴロ率":...}, "vsL": {...}}
    """
    out = {}
    for side_key in ("vsR", "vsL"):
        tot_pitches = swstr_sum = strike_sum = zone_sum = 0.0
        oswing_sum, oswing_count = 0.0, 0.0
        gb_sum, gb_count = 0.0, 0.0
        for ap in appearances:
            p = ap["player"]
            if not isinstance(p, dict):
                continue
            side = p.get("pitStatVsR" if side_key == "vsR" else "pitStatVsL")
            if not isinstance(side, dict):
                continue
            n = side.get("pitches", 0) or 0
            if n > 0:
                tot_pitches += n
                if side.get("swstr")  is not None: swstr_sum  += side["swstr"]  * n
                if side.get("strike") is not None: strike_sum += side["strike"] * n
                if side.get("zone")   is not None: zone_sum   += side["zone"]   * n
            if side.get("oSwing") is not None:
                ow = side.get("oz_n")
                if ow is not None:
                    oswing_sum += side["oSwing"] * ow
                    oswing_count += ow
            if side.get("gbpct") is not None:
                w = side.get("bip_known")
                if w is not None:
                    gn = side.get("gb_n")
                    gb_sum += gn if gn is not None else (side["gbpct"] / 100) * w
                    gb_count += w
        out[side_key] = {
            "空振り率":  round(swstr_sum / tot_pitches, 1) if tot_pitches > 0 else None,
            "ゾーン外スイング率": round(oswing_sum / oswing_count, 1) if oswing_count > 0 else None,
            "ストライク率": round(strike_sum / tot_pitches, 1) if tot_pitches > 0 else None,
            "ゾーン率":  round(zone_sum / tot_pitches, 1) if tot_pitches > 0 else None,
            "ゴロ率":   round(gb_sum / gb_count * 100, 1) if gb_count > 0 else None,
        }
    return out


# ==================================================
# Section 3. 球種別シーズン集計（_aggregateSeasonMix のポート）
# ==================================================

_CBS_FIELDS = ["c", "sw", "fo", "lo", "ba", "ou", "hi"]


def aggregate_season_mix(appearances: list[dict], mix_key: str = "mix") -> list[dict]:
    """
    mix_key: 'mix'（対戦打者を問わない合算） / 'mixVsR' / 'mixVsL'
    球種ごとに投球数加重平均で指標をまとめ、cbs（カウント別集計）もマージする。
    """
    km: dict[str, dict] = {}
    for ap in appearances:
        player = ap.get("player")
        if not isinstance(player, dict):
            continue
        if mix_key == "mixVsR":
            src = player.get("mixVsR") or player.get("mix") or []
        elif mix_key == "mixVsL":
            src = player.get("mixVsL") or player.get("mix") or []
        else:
            src = player.get("mix") or []
        if not isinstance(src, list):
            continue

        for m in src:
            if not isinstance(m, dict):
                continue
            key = m.get("key")
            if key not in km:
                km[key] = {
                    "name": m.get("name"), "key": key, "count": 0,
                    "swstr_sum": 0.0, "zone_sum": 0.0,
                    "oswing_sum": 0.0, "oswing_cnt": 0,
                    "vel_sum": 0.0, "vel_cnt": 0, "max_vel": 0.0,
                    "hits": 0, "hr": 0,
                    "strike_sum": 0.0, "strike_cnt": 0,
                    "gb_sum": 0.0, "gb_cnt": 0,
                    # MLB独自（NPBのmixにはキーが無いのでcnt=0のまま→Noneで出力）
                    "xwoba_sum": 0.0, "xwoba_cnt": 0,
                    "spin_sum": 0.0, "spin_cnt": 0,
                    "spinaxis_sum": 0.0, "spinaxis_cnt": 0,
                    "activespin_sum": 0.0, "activespin_cnt": 0,
                    "vaa_sum": 0.0, "vaa_cnt": 0,
                    "ivb_sum": 0.0, "ivb_cnt": 0,
                    "hb_sum": 0.0, "hb_cnt": 0,
                    "ext_sum": 0.0, "ext_cnt": 0,
                    "cbs_merged": defaultdict(lambda: {f: 0 for f in _CBS_FIELDS}),
                }
            k = km[key]
            count = m.get("count", 0) or 0
            k["count"] += count
            k["swstr_sum"] += (m.get("swstr") or 0) * count
            k["zone_sum"]  += (m.get("zone")  or 0) * count
            if m.get("oSwing") is not None:
                # ゾーン外スイング率の本来の分母は「その球種のゾーン外投球数」。
                # 「0」（その球種でゾーン外投球が無かった）は正当な値なのでそのまま使う。
                # oz_n自体が無い（まだ再生成していない古い形式の）試合は、投球数で代用すると
                # 単位が違う重みが混在して古い試合が不当に重く扱われるため、除外する。
                ow = m.get("oz_n")
                if ow is not None:
                    k["oswing_sum"] += (m["oSwing"] or 0) * ow
                    k["oswing_cnt"] += ow
            if m.get("vel"):
                k["vel_sum"] += m["vel"] * count
                k["vel_cnt"] += count
            if m.get("maxVel") and m["maxVel"] > k["max_vel"]:
                k["max_vel"] = m["maxVel"]
            k["hits"] += m.get("hits", 0) or 0
            k["hr"]   += m.get("hr", 0) or 0
            if m.get("strike") is not None:
                k["strike_sum"] += (m["strike"] or 0) * count
                k["strike_cnt"] += count
            if m.get("gbpct") is not None:
                # bip_knownが「0」（その球種で打球が1つも無かった）場合は正当な値なので
                # そのまま使う。bip_known自体が無い古い形式の試合は、投球数で代用すると
                # 単位が違う重みが混ざって古い試合が不当に重く扱われるため、除外する。
                w = m.get("bip_known")
                if w is not None:
                    gn = m.get("gb_n")
                    k["gb_sum"] += gn if gn is not None else ((m["gbpct"] or 0) / 100) * w
                    k["gb_cnt"] += w
            # MLB独自指標（投球数加重平均。NPBはこれらのキーが無いので蓄積されずcnt=0のまま）
            if m.get("xwoba") is not None:
                k["xwoba_sum"] += m["xwoba"] * count; k["xwoba_cnt"] += count
            if m.get("avgSpin") is not None:
                k["spin_sum"] += m["avgSpin"] * count; k["spin_cnt"] += count
            if m.get("spinAxis") is not None:
                k["spinaxis_sum"] += m["spinAxis"] * count; k["spinaxis_cnt"] += count
            if m.get("activeSpin") is not None:
                k["activespin_sum"] += m["activeSpin"] * count; k["activespin_cnt"] += count
            if m.get("vaa") is not None:
                k["vaa_sum"] += m["vaa"] * count; k["vaa_cnt"] += count
            if m.get("ivb") is not None:
                k["ivb_sum"] += m["ivb"] * count; k["ivb_cnt"] += count
            if m.get("hb") is not None:
                k["hb_sum"] += m["hb"] * count; k["hb_cnt"] += count
            if m.get("ext") is not None:
                k["ext_sum"] += m["ext"] * count; k["ext_cnt"] += count
            # cbs（カウント別）マージ
            cbs = m.get("cbs") or {}
            if isinstance(cbs, dict):
                for count_key, cv in cbs.items():
                    if not isinstance(cv, dict):
                        continue
                    cm = k["cbs_merged"][count_key]
                    for f in _CBS_FIELDS:
                        cm[f] += cv.get(f, 0) or 0

    merged = sorted(km.values(), key=lambda x: -x["count"])
    total = sum(m["count"] for m in merged)

    out = []
    for m in merged:
        out.append({
            "球種名": m["name"],
            "球種コード": m["key"],
            "投球数": m["count"],
            "投球割合%": round(m["count"] / total * 100, 1) if total > 0 else 0.0,
            "平均球速": round(m["vel_sum"] / m["vel_cnt"], 1) if m["vel_cnt"] > 0 else None,
            "最高球速": m["max_vel"] or None,
            "空振り率": round(m["swstr_sum"] / m["count"], 1) if m["count"] > 0 else 0.0,
            "ゾーン率": round(m["zone_sum"] / m["count"], 1) if m["count"] > 0 else 0.0,
            "ゾーン外スイング率": round(m["oswing_sum"] / m["oswing_cnt"], 1) if m["oswing_cnt"] > 0 else None,
            "ストライク率": round(m["strike_sum"] / m["strike_cnt"], 1) if m["strike_cnt"] > 0 else None,
            "GB%": round(m["gb_sum"] / m["gb_cnt"] * 100, 1) if m["gb_cnt"] > 0 else None,
            # ゴロ率・ゾーン外スイング率それぞれの本来の分母（打球数／ゾーン外投球数）の
            # 実数。この球種の複数試合分をすでに正しく加重した合計値なので、これをさらに
            # 別の場所（球種比較テーブルの「合計」行など）で加重平均する際の重みに使える。
            "GB数": round(m["gb_cnt"]) if m["gb_cnt"] else 0,
            "ゾーン外投球数": round(m["oswing_cnt"]) if m["oswing_cnt"] else 0,
            "H": m["hits"],
            "HR": m["hr"],
            # MLB独自（NPBはcnt=0のままなのでNone）
            "xwOBA": round(m["xwoba_sum"] / m["xwoba_cnt"], 3) if m["xwoba_cnt"] > 0 else None,
            "回転数": round(m["spin_sum"] / m["spin_cnt"]) if m["spin_cnt"] > 0 else None,
            "回転軸": round(m["spinaxis_sum"] / m["spinaxis_cnt"]) if m["spinaxis_cnt"] > 0 else None,
            "回転効率": round(m["activespin_sum"] / m["activespin_cnt"], 1) if m["activespin_cnt"] > 0 else None,
            "VAA": round(m["vaa_sum"] / m["vaa_cnt"], 1) if m["vaa_cnt"] > 0 else None,
            "縦変化量": round(m["ivb_sum"] / m["ivb_cnt"], 1) if m["ivb_cnt"] > 0 else None,
            "横変化量": round(m["hb_sum"] / m["hb_cnt"], 1) if m["hb_cnt"] > 0 else None,
            "Extension": round(m["ext_sum"] / m["ext_cnt"], 2) if m["ext_cnt"] > 0 else None,
            "_cbs": dict(m["cbs_merged"]),  # カウント別パターンシート生成用（出力シートには含めない）
        })
    return out


# ==================================================
# Section 4. コース分布（9分割ロジックのポート）
# ==================================================

_ZONE = 1.0
_EDGES9 = [-_ZONE, -_ZONE / 3, _ZONE / 3, _ZONE]


def _get_cell(v: float) -> int:
    for i in range(len(_EDGES9) - 1):
        if _EDGES9[i] <= v < _EDGES9[i + 1]:
            return i
    return 0 if v < _EDGES9[0] else len(_EDGES9) - 2


_OUTER = 1.67
_EDGES25 = [-_OUTER, -_ZONE, -_ZONE / 3, _ZONE / 3, _ZONE, _OUTER]

# pitcher-cards.html の hmBuildGridHtml と同じ5x5ビニング（ボール域込み）。
# row/col とも 0〜4。0/4がボール域、1〜3がストライクゾーン内3分割。
# 横位置は col=1側が内角、col=3側が外角（以前は逆になっていたため修正済み）。
_ROW_LABELS = {0: "ボール高め", 1: "高め", 2: "真ん中", 3: "低め", 4: "ボール低め"}
_COL_LABELS = {0: "ボール内角", 1: "内角", 2: "真ん中", 3: "外角", 4: "ボール外角"}


def _get_cell25(v: float) -> int:
    for i in range(len(_EDGES25) - 1):
        if _EDGES25[i] <= v < _EDGES25[i + 1]:
            return i
    return 0 if v < _EDGES25[0] else len(_EDGES25) - 2


def _pct(numer: int, denom: int) -> float | None:
    return round(numer / denom * 100, 1) if denom > 0 else None


def aggregate_course_grid25(locs_by_pitch: dict) -> list[dict]:
    """
    season_course_locs形式（{球種コード: {name, color, locsR, locsL}}）から、
    球種×対戦打者×25分割セルごとの球数・結果別割合を集計する。
    loc = [x, y, flag, inZone, rtype, isStrike]
      flag:  -1=見逃し/ボール, 0=スイング(空振り以外), 1=空振り
      rtype: 0=その他, 1=ゴロ, 2=フライ/ライナー, 3=単打, 4=二/三塁打, 5=本塁打

    サンプル閾値（プロンプト側のルールと一致させる）:
      - ゾーン内セル（row・colとも1〜3）: 15球以上
      - ボール域セル（row・colどちらかが0か4）: 10球以上
    「無駄球/釣り球」判定はボール域セルかつ閾値以上の場合のみ付与する:
      - 見逃し率75%以上 → "無駄球"
      - 空振り率20%以上 → "釣り球"
      - それ以外 → None
    """
    rows_out: list[dict] = []
    for pt_code, obj in (locs_by_pitch or {}).items():
        pt_name = obj.get("name") or pt_code
        for side_key, side_label in (("locsR", "対右打者"), ("locsL", "対左打者")):
            cells: dict[tuple[int, int], dict] = {}
            for loc in (obj.get(side_key) or []):
                if not isinstance(loc, list) or len(loc) < 5:
                    continue
                x, y, flag, _in_zone, rtype = loc[0], loc[1], loc[2], loc[3], loc[4]
                col = _get_cell25(x)
                row = _get_cell25(-y)
                key = (row, col)
                if key not in cells:
                    cells[key] = {"total": 0, "take": 0, "whiff": 0,
                                  "ground": 0, "fly": 0, "single": 0, "xbh": 0, "hr": 0}
                c = cells[key]
                c["total"] += 1
                if flag == -1:
                    c["take"] += 1
                elif flag == 1:
                    c["whiff"] += 1
                if rtype == 1:
                    c["ground"] += 1
                elif rtype == 2:
                    c["fly"] += 1
                elif rtype == 3:
                    c["single"] += 1
                elif rtype in (4, 5):
                    c["xbh"] += 1
                    if rtype == 5:
                        c["hr"] += 1

            for row in range(5):
                for col in range(5):
                    c = cells.get((row, col))
                    total = c["total"] if c else 0
                    is_ball_zone = row in (0, 4) or col in (0, 4)
                    threshold = 10 if is_ball_zone else 15
                    sample_ok = total >= threshold
                    take_pct = _pct(c["take"], total) if c else None
                    whiff_pct = _pct(c["whiff"], total) if c else None
                    waste_flag = None
                    if sample_ok and is_ball_zone:
                        if take_pct is not None and take_pct >= 75:
                            waste_flag = "無駄球"
                        elif whiff_pct is not None and whiff_pct >= 20:
                            waste_flag = "釣り球"
                    rows_out.append({
                        "球種名": pt_name,
                        "対戦打者": side_label,
                        "縦位置": _ROW_LABELS[row],
                        "横位置": _COL_LABELS[col],
                        "ボール域": is_ball_zone,
                        "球数": total,
                        "見逃し率%": take_pct,
                        "空振り率%": whiff_pct,
                        "ゴロ率%": _pct(c["ground"], total) if c else None,
                        "フライ率%": _pct(c["fly"], total) if c else None,
                        "単打率%": _pct(c["single"], total) if c else None,
                        "長打率%": _pct(c["xbh"], total) if c else None,
                        "本塁打数": c["hr"] if c else 0,
                        "サンプル十分": sample_ok,
                        "判定": waste_flag,
                    })
    return rows_out


def aggregate_course_distribution(appearances: list[dict]) -> dict:
    """
    対右打者(R) / 対左打者(L) それぞれについて、ゾーン内投球を3x3(9マス)に分類し
    件数・割合を返す。index.htmlの9分割ロジック（投手視点、ゾーン内のみ分母）と同じ。
    戻り値: {"R": {"total_in_zone": n, "cells": [[cnt,...]x3]}, "L": {...}}
    """
    result = {"R": [[0, 0, 0] for _ in range(3)], "L": [[0, 0, 0] for _ in range(3)]}
    totals = {"R": 0, "L": 0}

    for ap in appearances:
        player = ap.get("player")
        if not isinstance(player, dict):
            continue
        for side_key, mix_key in (("R", "mixVsR"), ("L", "mixVsL")):
            for m in (player.get(mix_key) or []):
                if not isinstance(m, dict):
                    continue
                for loc in (m.get("locs") or []):
                    # loc: [x, y, resultCode, inZoneFlag, ...]
                    if not isinstance(loc, list) or len(loc) < 4 or loc[3] != 1:
                        continue
                    x, y = loc[0], loc[1]
                    col = _get_cell(x)
                    row = _get_cell(-y)
                    result[side_key][row][col] += 1
                    totals[side_key] += 1

    out = {}
    for side_key in ("R", "L"):
        cells = result[side_key]
        total = totals[side_key]
        rows = []
        for r in range(3):
            for c in range(3):
                cnt = cells[r][c]
                rows.append({
                    "ゾーン番号": r * 3 + c + 1,  # 1〜9（左上→右下）
                    "球数": cnt,
                    "割合%": round(cnt / total * 100, 1) if total > 0 else 0.0,
                })
        out[side_key] = {"total_in_zone": total, "rows": rows}
    return out


# ==================================================
# Section 5. カウント別パターン（cbsマージ結果を整形）
# ==================================================

# ==================================================
# Section 3.5. 球種別カラースケール（本家ダッシュボードの getScaleColor と同じ仕組み）
#   役割（先発/中継ぎ）× 球種コード ごとに、全投手・全登板の球種別データから
#   p15/p30/p40/p60/p70/p85 パーセンタイルを算出し、値がどの階層に入るかを判定する。
# ==================================================

_COLOR_SCALE_METRICS = [
    "swstr", "oSwing", "strike", "zone", "gbpct", "vel", "maxVel",
    # MLB独自（Statcast由来。NPBデータには存在しないため自動的にNoneスキップされる）
    "avgSpin", "spinAxis", "activeSpin", "vaa", "ivb", "hb", "ext", "xwoba",
]
_COLOR_SCALE_DIR = {
    "swstr": True, "oSwing": True, "strike": True, "zone": True, "gbpct": True,
    "vel": True, "maxVel": True,
    # MLB独自。index.html の COL_SCALE_DIR と同じ方向定義。
    "avgSpin": True, "spinAxis": False, "activeSpin": True, "vaa": False,
    "ivb": True, "hb": True, "ext": True, "xwoba": False,
}
# True = 値が高いほど良い（ティール方向）。dashboardのCOL_SCALE_DIRと同じ定義。


def _pitcher_role(player: dict) -> str:
    role = (player.get("role") or "").strip()
    return "starter" if role == "先発" else "reliever"


def build_pitch_color_scale_stats(all_data: dict) -> dict:
    """
    dashboard.html の buildPitchColorScaleStats() のポート。
    {role: {pitch_key: {metric: {p15,p30,p40,p60,p70,p85}}}} を返す。
    """
    by_role: dict[str, dict[str, dict[str, list]]] = {"starter": {}, "reliever": {}}

    for date, games in all_data.items():
        if date == "highlights" or date.startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict):
                continue
            pitchers = g.get("pitchers")
            if not isinstance(pitchers, dict):
                continue
            for side in ("home", "away"):
                for p in (pitchers.get(side) or []):
                    if not isinstance(p, dict):
                        continue
                    role_key = _pitcher_role(p)
                    for m in (p.get("mix") or []):
                        if not isinstance(m, dict):
                            continue
                        key = m.get("key")
                        if key is None:
                            continue
                        bucket = by_role[role_key].setdefault(
                            key, {metric: [] for metric in _COLOR_SCALE_METRICS}
                        )
                        for metric in _COLOR_SCALE_METRICS:
                            v = m.get(metric)
                            if v is not None:
                                bucket[metric].append(float(v))

    def calc_stats(by_key: dict) -> dict:
        result = {}
        for pitch_key, val_map in by_key.items():
            result[pitch_key] = {}
            for metric, arr in val_map.items():
                if not arr:
                    result[pitch_key][metric] = {"p15": 0, "p30": 0, "p40": 0, "p60": 100, "p70": 100, "p85": 100}
                    continue
                arr = sorted(arr)
                n = len(arr)
                def pct(p):
                    return arr[min(n - 1, int(n * p / 100))]
                result[pitch_key][metric] = {
                    "p15": pct(15), "p30": pct(30), "p40": pct(40),
                    "p60": pct(60), "p70": pct(70), "p85": pct(85),
                }
        return result

    return {
        "starter": calc_stats(by_role["starter"]),
        "reliever": calc_stats(by_role["reliever"]),
    }


def get_scale_tier(metric: str, value, stats_for_pitch: dict) -> str | None:
    """getScaleColor() のランク判定部分のポート。'top15'〜'bot85' または None（色なし）"""
    if value is None:
        return None
    direction = _COLOR_SCALE_DIR.get(metric)
    if direction is None:
        return None
    s = stats_for_pitch.get(metric)
    if not s:
        return None
    v = float(value)
    if direction:
        if v >= s["p85"]: return "top15"
        if v >= s["p70"]: return "top30"
        if v >= s["p60"]: return "top40"
        if v >= s["p40"]: return "mid"
        if v >= s["p30"]: return "bot60"
        if v >= s["p15"]: return "bot70"
        return "bot85"
    else:
        if v <= s["p15"]: return "top15"
        if v <= s["p30"]: return "top30"
        if v <= s["p40"]: return "top40"
        if v <= s["p60"]: return "mid"
        if v <= s["p70"]: return "bot60"
        if v <= s["p85"]: return "bot70"
        return "bot85"


def annotate_mix_rows_with_tiers(mix_rows: list[dict], role_key: str, pitch_scale_stats: dict) -> None:
    """
    aggregate_season_mix() が返す行（辞書）に、色付け用の rank_tier を直接追加する（in-place）。
    キー対応: 空振り率→swstr, ゾーン外スイング率→oSwing, ストライク率→strike, ゾーン率→zone, GB%→gbpct
    """
    stats_for_role = pitch_scale_stats.get(role_key, {})
    field_map = {
        "空振り率": "swstr",
        "ゾーン外スイング率": "oSwing",
        "ストライク率": "strike",
        "ゾーン率": "zone",
        "GB%": "gbpct",
    }
    for row in mix_rows:
        pitch_key = row.get("球種コード")
        stats_for_pitch = stats_for_role.get(pitch_key, {})
        tiers = {}
        for field_jp, metric in field_map.items():
            tiers[field_jp] = get_scale_tier(metric, row.get(field_jp), stats_for_pitch)
        row["_rank_tier"] = tiers


def merge_lr_split(mix_all: list[dict], mix_vs_r: list[dict], mix_vs_l: list[dict]) -> list[dict]:
    """
    球種ごとの対右/対左スタッツを、全体集計(mix_all)の行にマージする。
    球種評価で「この球種は対左打者にどうか」まで言及できるようにするための追加情報。

    空振り率・ゴロ率・H・HRに加え、ゾーン外スイング率・ストライク率・ゾーン率も
    対右/対左それぞれ追加する（compute_pitch_rankings() で対右/対左の順位を
    算出するために必要な指標一式を揃えるため）。
    """
    r_by_key = {m["球種コード"]: m for m in mix_vs_r}
    l_by_key = {m["球種コード"]: m for m in mix_vs_l}

    out = []
    for m in mix_all:
        row = dict(m)
        r = r_by_key.get(m["球種コード"])
        l = l_by_key.get(m["球種コード"])
        row["対右_投球数"] = r["投球数"] if r else 0
        row["対右_空振り率"] = r["空振り率"] if r else None
        row["対右_ゾーン外スイング率"] = r["ゾーン外スイング率"] if r else None
        row["対右_ストライク率"] = r["ストライク率"] if r else None
        row["対右_ゾーン率"] = r["ゾーン率"] if r else None
        row["対右_ゴロ率"] = r["GB%"] if r else None
        row["対右_H"] = r["H"] if r else 0
        row["対右_HR"] = r["HR"] if r else 0
        row["対左_投球数"] = l["投球数"] if l else 0
        row["対左_空振り率"] = l["空振り率"] if l else None
        row["対左_ゾーン外スイング率"] = l["ゾーン外スイング率"] if l else None
        row["対左_ストライク率"] = l["ストライク率"] if l else None
        row["対左_ゾーン率"] = l["ゾーン率"] if l else None
        row["対左_ゴロ率"] = l["GB%"] if l else None
        row["対左_H"] = l["H"] if l else 0
        row["対左_HR"] = l["HR"] if l else 0
        out.append(row)
    return out


def build_count_pattern_rows(player_name: str, season_mix: list[dict]) -> list[dict]:
    """season_mix の各球種が持つ _cbs を、カウント別パターンシート用の行に変換"""
    rows = []
    for m in season_mix:
        for count_key, cv in (m.get("_cbs") or {}).items():
            c = cv.get("c", 0) or 0
            if c == 0:
                continue
            rows.append({
                "選手名": player_name,
                "球種名": m["球種名"],
                "カウント": count_key,  # 例: "0-0", "0-2", "1-2" など
                "球数": c,
                "空振り": cv.get("sw", 0),
                "ファウル": cv.get("fo", 0),
                "見逃し": cv.get("lo", 0),
                "ボール": cv.get("ba", 0),
                "凡打": cv.get("ou", 0),
                "被安打": cv.get("hi", 0),
                "空振り率%": round(cv.get("sw", 0) / c * 100, 1),
                "ファウル率%": round(cv.get("fo", 0) / c * 100, 1),
            })
    return rows


# ==================================================
# Section 6. メイン: 全投手を対象に4シート分のDataFrameを組み立ててxlsx出力
# ==================================================

# ==================================================
# Section 5.5. 試合ログ（1試合=1行。カードの「試合成績」テーブル用）
# ==================================================

def _single_game_mix_rows(mix_list) -> list[dict]:
    """1試合ぶんのmix配列を、カード表示用の球種行形式に整形（ゾーン率・ゴロ率・球速・MLB独自指標も含む）"""
    if not isinstance(mix_list, list):
        return []
    valid = [m for m in mix_list if isinstance(m, dict)]
    total = sum((m.get("count") or 0) for m in valid)
    rows = []
    for m in sorted(valid, key=lambda x: -(x.get("count") or 0)):
        count = m.get("count") or 0
        rows.append({
            "name": m.get("name"),
            "count": count,
            "pct": round(count / total * 100, 1) if total > 0 else 0.0,
            "swstr_pct": m.get("swstr"),
            "chase_pct": m.get("oSwing"),
            "strike_pct": m.get("strike"),
            "zone_pct": m.get("zone"),
            "gb_pct": m.get("gbpct"),
            # gb_n: html側のwavg()フォールバックで「重み（分母）」として使われるフィールド。
            # 分子（ゴロの個数）ではなく、必ずbip_known（打球数）を入れる。
            "gb_n": m.get("bip_known"),
            "oz_n": m.get("oz_n"),
            "avg_vel": m.get("vel"),
            "max_vel": m.get("maxVel"),
            # MLB独自（Statcast由来。NPBのmixオブジェクトにはキー自体が無いのでNoneになる）
            "xwoba": m.get("xwoba"),
            "avg_spin": m.get("avgSpin"),
            "spin_axis": m.get("spinAxis"),
            "active_spin": m.get("activeSpin"),
            "vaa": m.get("vaa"),
            "ivb": m.get("ivb"),
            "hb": m.get("hb"),
            "ext": m.get("ext"),
        })
    return rows


def _single_game_pitch_tiers(mix_rows: list[dict], role_key: str, pitch_scale_stats: dict, key_lookup: dict) -> None:
    """1試合ぶんのmix行に、season版と同じロジックでランクタグを付与する（in-place）"""
    stats_for_role = pitch_scale_stats.get(role_key, {})
    for row in mix_rows:
        pitch_key = key_lookup.get(row.get("name"))
        stats_for_pitch = stats_for_role.get(pitch_key, {}) if pitch_key else {}
        row["空振り率_ランク"] = get_scale_tier("swstr", row.get("swstr_pct"), stats_for_pitch)
        row["ゾーン外スイング率_ランク"] = get_scale_tier("oSwing", row.get("chase_pct"), stats_for_pitch)
        row["ストライク率_ランク"] = get_scale_tier("strike", row.get("strike_pct"), stats_for_pitch)
        row["ゾーン率_ランク"] = get_scale_tier("zone", row.get("zone_pct"), stats_for_pitch)
        row["GB%_ランク"] = get_scale_tier("gbpct", row.get("gb_pct"), stats_for_pitch)
        row["平均球速_ランク"] = get_scale_tier("vel", row.get("avg_vel"), stats_for_pitch)
        row["最高球速_ランク"] = get_scale_tier("maxVel", row.get("max_vel"), stats_for_pitch)
        # MLB独自（NPBはvalueがNoneなのでget_scale_tierがNoneを返して自動的に色なしになる）
        row["回転数_ランク"] = get_scale_tier("avgSpin", row.get("avg_spin"), stats_for_pitch)
        row["回転軸_ランク"] = get_scale_tier("spinAxis", row.get("spin_axis"), stats_for_pitch)
        row["回転効率_ランク"] = get_scale_tier("activeSpin", row.get("active_spin"), stats_for_pitch)
        row["VAA_ランク"] = get_scale_tier("vaa", row.get("vaa"), stats_for_pitch)
        row["縦変化量_ランク"] = get_scale_tier("ivb", row.get("ivb"), stats_for_pitch)
        row["横変化量_ランク"] = get_scale_tier("hb", row.get("hb"), stats_for_pitch)
        row["Extension_ランク"] = get_scale_tier("ext", row.get("ext"), stats_for_pitch)
        row["xwOBA_ランク"] = get_scale_tier("xwoba", row.get("xwoba"), stats_for_pitch)


_PITCH_COLOR_MAP = {
    "FF": "#3B82F6", "SL": "#F59E0B", "CU": "#10B981", "FK": "#F87171",
    "CH": "#A78BFA", "SI": "#22D3EE", "CT": "#FB923C", "SP": "#2DD4BF",
    "FS": "#2DD4BF", "SH": "#F472B6", "ST": "#2DD4BF",
}


def _extract_locs_from_appearances(appearances: list[dict]) -> dict:
    """
    複数登板ぶんのappearancesから、球種コードごとに {name, color, locsR, locsL} を集める。
    本家ダッシュボードの _drawZoneHeatmapMulti に渡す pitchLocsMap と同じ形式。
    locは [x, y, flag, inZone, rtype, isStrike] の6要素に切り詰める（余分なフィールドは捨てる）。
    """
    out: dict[str, dict] = {}
    for ap in appearances:
        player = ap.get("player")
        if not isinstance(player, dict):
            continue
        for side_key, mix_key in (("locsR", "mixVsR"), ("locsL", "mixVsL")):
            for m in (player.get(mix_key) or []):
                if not isinstance(m, dict):
                    continue
                key = m.get("key")
                if key is None:
                    continue
                if key not in out:
                    out[key] = {
                        "name": m.get("name"),
                        "color": _PITCH_COLOR_MAP.get(key, "#94A3B8"),
                        "locsR": [], "locsL": [],
                    }
                for loc in (m.get("locs") or []):
                    if isinstance(loc, list) and len(loc) >= 6:
                        out[key][side_key].append(loc[:6])
    return out


def build_season_course_locs(appearances: list[dict]) -> dict:
    """シーズン全体ぶんの球種別コース座標（pitchLocsMap形式）"""
    return _extract_locs_from_appearances(appearances)


def _single_game_course_locs(ap: dict) -> dict:
    """1試合ぶんの球種別コース座標（pitchLocsMap形式）"""
    return _extract_locs_from_appearances([ap])


def _course_result_to_detail(course_result: dict) -> dict:
    """aggregate_course_distribution()の戻り値を、カード表示用の対右/対左9セル pct+count 形式に整形する"""
    out = {}
    for side_key in ("R", "L"):
        total = course_result[side_key]["total_in_zone"]
        cells = [{"pct": r["割合%"], "count": r["球数"]} for r in course_result[side_key]["rows"]]
        out["vsR" if side_key == "R" else "vsL"] = {"total": total, "cells": cells}
    return out


def _single_game_course_detail(ap: dict) -> dict:
    """1試合ぶんのappearanceから、カード表示用のコース分布（対右/対左 9セル pct+count）を作る"""
    return _course_result_to_detail(aggregate_course_distribution([ap]))


def build_season_pitch_detail(season_mix_all, season_mix_vs_r, season_mix_vs_l,
                               role_key: str, pitch_scale_stats: dict, key_lookup: dict) -> dict:
    """シーズン全体の球種詳細を、試合ごとの球種詳細と同じ {all, vsR, vsL} 形式で作る"""
    def _fmt(mix_list):
        rows = [{
            "name": m["球種名"],
            "count": m["投球数"],
            "pct": m["投球割合%"],
            "swstr_pct": m["空振り率"],
            "chase_pct": m["ゾーン外スイング率"],
            "strike_pct": m["ストライク率"],
            "zone_pct": m["ゾーン率"],
            "gb_pct": m.get("GB%"),
            # 合計行の加重平均用（gb_pctはGB数、chase_pctはゾーン外投球数が本来の重み）
            "gb_n": m.get("GB数"),
            "oz_n": m.get("ゾーン外投球数"),
            "avg_vel": m.get("平均球速"),
            "max_vel": m.get("最高球速"),
            # MLB独自（NPBはaggregate_season_mix側でNoneになる）
            "xwoba": m.get("xwOBA"),
            "avg_spin": m.get("回転数"),
            "spin_axis": m.get("回転軸"),
            "active_spin": m.get("回転効率"),
            "vaa": m.get("VAA"),
            "ivb": m.get("縦変化量"),
            "hb": m.get("横変化量"),
            "ext": m.get("Extension"),
        } for m in mix_list]
        _single_game_pitch_tiers(rows, role_key, pitch_scale_stats, key_lookup)
        return rows
    return {"all": _fmt(season_mix_all), "vsR": _fmt(season_mix_vs_r), "vsL": _fmt(season_mix_vs_l)}


def build_game_log_rows(name: str, appearances: list[dict], role_key: str, pitch_scale_stats: dict) -> list[dict]:
    """投手1人分の試合ログ（1試合=1行）。カードJSONのgame_logにそのまま使える形。"""
    rows = []
    # 球種名→球種コードの対応（この投手のseason mixから作る。ランク付けにコードが必要なため）
    season_mix_all = aggregate_season_mix(appearances, "mix")
    key_lookup = {m["球種名"]: m["球種コード"] for m in season_mix_all}

    for ap in sorted(appearances, key=lambda a: a["date"], reverse=True):
        p = ap["player"]
        if not isinstance(p, dict):
            continue
        game = ap.get("game") or {}
        side = ap.get("side")
        opponent_side = "away" if side == "home" else "home"
        opponent = game.get(opponent_side) if isinstance(game, dict) else None

        pitch_detail_all = _single_game_mix_rows(p.get("mix"))
        pitch_detail_vsr = _single_game_mix_rows(p.get("mixVsR"))
        pitch_detail_vsl = _single_game_mix_rows(p.get("mixVsL"))
        _single_game_pitch_tiers(pitch_detail_all, role_key, pitch_scale_stats, key_lookup)
        _single_game_pitch_tiers(pitch_detail_vsr, role_key, pitch_scale_stats, key_lookup)
        _single_game_pitch_tiers(pitch_detail_vsl, role_key, pitch_scale_stats, key_lookup)

        tbf = p.get("tbf") or ((p.get("k") or 0) + (p.get("bb") or 0) + (p.get("h") or 0))
        k_pct = round((p.get("k") or 0) / tbf * 100, 1) if tbf else None
        bb_pct = round((p.get("bb") or 0) / tbf * 100, 1) if tbf else None
        kbb_pct = round(((p.get("k") or 0) - (p.get("bb") or 0)) / tbf * 100, 1) if tbf else None

        rows.append({
            "date": ap["date"],
            "opponent": f"vs {opponent}" if opponent else None,
            "result": p.get("result") or "ND",
            "swstr_pct": p.get("swstr"),
            "chase_pct": p.get("oSwing"),
            "strike_pct": p.get("strike"),
            "zone_pct": p.get("zone"),
            "kbb_pct": kbb_pct,
            "k_pct": k_pct,
            "bb_pct": bb_pct,
            "gb_pct": p.get("gbpct"),
            "ip": p.get("ip"),
            "pitches": p.get("pitches"),
            "k": p.get("k"),
            "bb": p.get("bb"),
            "h": p.get("h"),
            "er": p.get("er"),
            "pitch_detail": {"all": pitch_detail_all, "vsR": pitch_detail_vsr, "vsL": pitch_detail_vsl},
            "course_detail": _single_game_course_detail(ap),
            "course_locs": _single_game_course_locs(ap),
        })
    return rows


def compute_rankings(season_rows: list[dict],
                      rank_min_ip: dict[str, float] | None = None) -> dict:
    """
    シーズン集計の投手分から、防御率・K-BB%・K%・BB%・ゴロ率の順位を算出する。

    役割（先発/中継ぎ）ごとに母集団を分けて順位を出す（「先発の中での順位」「中継ぎの中での
    順位」になる。全投手混合の順位は出さない）。
    さらに、役割ごとの資格投球回（rank_min_ip）未満の投手は順位の母集団そのものから除外する
    （xlsx自体への掲載可否＝min_ip引数とは別軸。min_ipで出力対象になっていても、
    投球回が少なければ順位は付与されずrankingsに現れない＝カード側では非表示になる）。
    デフォルトの資格投球回は、球種別詳細側の順位（RANK_MIN_IP）と同じ
    先発=投球回15回以上、中継ぎ=投球回10回以上に揃えている。

    戻り値: {選手名: {"era":{"rank":n,"total":m}, "k_bb_pct":{...}, "k_pct":{...}, "bb_pct":{...}, "gb_pct":{...}}}
    （資格投球回に満たない選手のエントリは空辞書 {} のまま＝どの指標も順位が付かない）
    """
    if rank_min_ip is None:
        rank_min_ip = {"先発": 15.0, "中継ぎ": 10.0}

    specs = [
        ("防御率", "era", False),      # 低いほど良い
        ("K-BB%", "k_bb_pct", True),   # 高いほど良い
        ("K%", "k_pct", True),         # 高いほど良い
        ("BB%", "bb_pct", False),      # 低いほど良い
        ("ゴロ率", "gb_pct", True),    # 高いほど良い
    ]
    result = {row["選手名"]: {} for row in season_rows}

    for role, threshold in rank_min_ip.items():
        role_rows = [
            row for row in season_rows
            if row.get("役割") == role and (_ip_to_outs(row.get("投球回")) / 3) >= threshold
        ]
        for jp_key, out_key, higher_is_better in specs:
            valid = [(row["選手名"], row[jp_key]) for row in role_rows if row.get(jp_key) is not None]
            valid.sort(key=lambda x: -x[1] if higher_is_better else x[1])
            total = len(valid)
            for rank, (name, _) in enumerate(valid, start=1):
                result[name][out_key] = {"rank": rank, "total": total, "role": role}
    return result


# ==================================================
# Section 6.5. 球種別順位（全体/対右/対左）
#   xlsx「球種別詳細」シート向け。旧tier方式（_ランク列、pitch_scale_stats／
#   get_scale_tier）を置き換えるもの。
#   ※ pitch_scale_stats を使ったtier方式は、numeric_json_dir（ダッシュボードの
#     カードJSON。season_pitch_detail・game_logの色分け表示用）では引き続き
#     使うため、build_pitch_color_scale_stats/get_scale_tier/
#     annotate_mix_rows_with_tiers 自体は削除しない。
# ==================================================

# 順位算出の対象条件。自動実行パイプラインのためCLI引数にはせず定数で固定する。
RANK_MIN_IP = {"先発": 15.0, "中継ぎ": 10.0}      # 全体側の順位: 役割別の資格投球回
RANK_MIN_PITCHES_VS_HAND = 20                          # 対右/対左側の順位: 球種ごとの対戦数条件（役割共通）

# 順位を出す指標。(全体側のフィールド名, 高いほど良いか)
# 対右/対左側は同名の接頭辞（対右_/対左_）を付けたフィールドを見る。
# GB%だけ対右/対左側は「ゴロ率」という列名でマージされている点に注意（merge_lr_split参照）。
_PITCH_RANK_METRICS_ALL = [
    ("空振り率", True), ("ゾーン外スイング率", True),
    ("ストライク率", True), ("ゾーン率", True), ("GB%", True),
]
_PITCH_RANK_METRICS_HAND = [
    ("空振り率", True), ("ゾーン外スイング率", True),
    ("ストライク率", True), ("ゾーン率", True), ("ゴロ率", True),
]


def compute_pitch_rankings(mix_rows: list[dict], role_map: dict[str, str], ip_map: dict[str, str]) -> None:
    """
    mix_rows（球種別詳細シートの全選手ぶんの行）に、全体/対右/対左それぞれの
    「{指標}_順位」「{指標}_順位_母数」列をin-placeで追加する。

    母集団は 役割（先発/中継ぎ）× 球種コード ごとに分ける
    （先発の中での順位、中継ぎの中での順位。役割混合の順位は出さない）。

    _順位:
        以下の条件を満たす投手だけの母集団内で計算する順位。満たさない投手はnull。
          - 役割別の資格投球回（RANK_MIN_IP、シーズン通算投球回で判定）以上
          - 対右/対左側はさらに、その球種の対右_投球数/対左_投球数が
            RANK_MIN_PITCHES_VS_HAND 以上
    _順位_母数:
        上記の条件でフィルタせず、同じ役割内でその指標の値を持つ投手全員を母数にする
        （投球回や対戦数の条件を満たすかどうかは問わない）。
    """

    def ip_float(name: str) -> float:
        return _ip_to_outs(ip_map.get(name)) / 3

    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in mix_rows:
        name = row.get("選手名")
        role = role_map.get(name)
        pitch_key = row.get("球種コード")
        if role is None or pitch_key is None:
            continue
        groups[(role, pitch_key)].append(row)

    def assign(rows: list[dict], value_field: str, min_ip: float | None,
               higher_is_better: bool, count_field: str | None) -> None:
        all_pairs = [(r, r.get(value_field)) for r in rows if r.get(value_field) is not None]
        total_all = len(all_pairs)

        qualified = []
        for r, v in all_pairs:
            if min_ip is None or ip_float(r["選手名"]) < min_ip:
                continue
            if count_field is not None and (r.get(count_field) or 0) < RANK_MIN_PITCHES_VS_HAND:
                continue
            qualified.append((r, v))
        qualified.sort(key=lambda x: -x[1] if higher_is_better else x[1])

        rank_field = f"{value_field}_順位"
        total_field = f"{value_field}_順位_母数"
        for r in rows:
            r[rank_field] = None
            r[total_field] = total_all
        for i, (r, _v) in enumerate(qualified, start=1):
            r[rank_field] = i

    for (role, _pitch_key), rows in groups.items():
        min_ip = RANK_MIN_IP.get(role)
        for field, higher in _PITCH_RANK_METRICS_ALL:
            assign(rows, field, min_ip, higher, count_field=None)
        for side_prefix, count_field in (("対右", "対右_投球数"), ("対左", "対左_投球数")):
            for field, higher in _PITCH_RANK_METRICS_HAND:
                assign(rows, f"{side_prefix}_{field}", min_ip, higher, count_field=count_field)


def determine_pitcher_role(appearances: list[dict]) -> str:
    """登板ごとのroleの最頻値で、その投手のシーズンを通した役割を決める"""
    from collections import Counter
    roles = [ap["player"].get("role") for ap in appearances if isinstance(ap.get("player"), dict)]
    roles = [r for r in roles if r]
    if not roles:
        return "starter"
    most_common = Counter(roles).most_common(1)[0][0]
    return "starter" if most_common == "先発" else "reliever"


def classify_pitcher_categories(card: dict) -> list:
    """
    球種構成・ゴロ/フライ傾向から、選手一覧で絞り込みに使えるカテゴリタグを機械的に組み立てる。
    複数タグが同時に付き得る（例：「フライピッチャー」＋「ストレートピッチャー」＋「スライダーピッチャー」）。
    """
    cats = []
    gb = card.get("gb_pct")
    if gb is not None:
        if gb >= 50:
            cats.append("ゴロピッチャー")
        elif gb <= 38:
            cats.append("フライピッチャー")

    pitches = ((card.get("season_pitch_detail") or {}).get("all")) or []
    for p in pitches:
        name = p.get("name")
        pct = p.get("pct")
        # 投球割合20%以上を「持ち味の球種」とみなす（1人の投手が複数該当してOK）
        if name and pct is not None and pct >= 20:
            cats.append(f"{name}ピッチャー")
    return cats


def export_llm_input_xlsx(games_json_dir: str, out_path: str, min_ip: float = 0.0,
                           target_names: list[str] | None = None,
                           numeric_json_dir: str | None = None) -> str:
    """
    numeric_json_dir を指定すると、xlsxに加えて選手ごとの数値データJSON
    （pitcher_cards_numeric/{選手ID}.json）と選手一覧 index.json も書き出す。
    これらはLLMの解釈テキストを含まない「数値だけ」のデータで、pitcher-cards.html側が
    LLM出力（バッチJSON）とブラウザ内でマージして使う。
    """
    all_data = load_daily_games(games_json_dir)
    names = set(target_names) if target_names else build_all_pitcher_names(all_data)

    print("  球種別カラースケール（パーセンタイル）を算出中...")
    pitch_scale_stats = build_pitch_color_scale_stats(all_data)

    season_rows, mix_rows, course_rows, count_rows, gamelog_rows = [], [], [], [], []
    course_grid_rows = []
    numeric_cards: dict[str, dict] = {}  # {選手名: numeric card dict}（xlsx書き出し後、rankings付与してからJSON化）

    for name in sorted(names):
        try:
            appearances = build_appearances(all_data, name)
            if not appearances:
                continue

            season = calc_season_stats(appearances)
            season["選手名"] = name  # 表記ゆれ正規化後の名前で統一
            season_by_hand = calc_season_stats_by_hand(appearances)

            # 投球回フィルタ（例: 10回以上登板した投手のみ対象）
            ip_num = _ip_to_outs(season["投球回"]) / 3
            if ip_num < min_ip:
                continue

            role_key = determine_pitcher_role(appearances)
            season["役割"] = "先発" if role_key == "starter" else "中継ぎ"

            # 直近の登板からチーム名を推定（home/awayどちら側だったかで判定）
            last_ap = appearances[-1]
            last_game = last_ap.get("game") or {}
            team = last_game.get(last_ap.get("side")) if isinstance(last_game, dict) else None
            season["所属チーム"] = team

            season_rows.append(season)

            season_mix_all = aggregate_season_mix(appearances, "mix")
            season_mix_vs_r = aggregate_season_mix(appearances, "mixVsR")
            season_mix_vs_l = aggregate_season_mix(appearances, "mixVsL")
            season_mix_merged = merge_lr_split(season_mix_all, season_mix_vs_r, season_mix_vs_l)
            annotate_mix_rows_with_tiers(season_mix_merged, role_key, pitch_scale_stats)

            pitch_numeric_rows = []
            for m in season_mix_merged:
                tiers = m.get("_rank_tier", {})
                # xlsx（球種別詳細シート）側の _ランク（tier）列は廃止。
                # 全体/対右/対左の _順位・_順位_母数 は、全選手ぶん集まった後に
                # compute_pitch_rankings() でまとめて付与する。
                row = {"選手名": name, **{k: v for k, v in m.items() if k not in ("_cbs", "_rank_tier")}}
                mix_rows.append(row)
                pitch_numeric_rows.append({
                    "name": m.get("球種名"),
                    "count": m.get("投球数"),
                    "pct": m.get("投球割合%"),
                    "swstr_pct_rank": tiers.get("空振り率"),
                    "chase_pct_rank": tiers.get("ゾーン外スイング率"),
                    "strike_pct_rank": tiers.get("ストライク率"),
                })

            count_rows.extend(build_count_pattern_rows(name, season_mix_all))

            course = aggregate_course_distribution(appearances)
            for side_key, side_label in (("R", "対右打者"), ("L", "対左打者")):
                for row in course[side_key]["rows"]:
                    course_rows.append({
                        "選手名": name,
                        "対戦打者": side_label,
                        "ゾーン内総数": course[side_key]["total_in_zone"],
                        **row,
                    })

            # シーズン全体の球種詳細（全体/対右/対左）とコース分布（試合ごとの行と同じ形式）
            key_lookup = {m["球種名"]: m["球種コード"] for m in season_mix_all}
            season_pitch_detail = build_season_pitch_detail(
                season_mix_all, season_mix_vs_r, season_mix_vs_l, role_key, pitch_scale_stats, key_lookup
            )
            season_course_detail = _course_result_to_detail(course)
            season_course_locs = build_season_course_locs(appearances)

            for row in aggregate_course_grid25(season_course_locs):
                course_grid_rows.append({"選手名": name, **row})

            # 試合ログ（1試合=1行。pitch_detail/course_detailはネスト構造なのでJSON文字列として保持）
            game_log_dicts = build_game_log_rows(name, appearances, role_key, pitch_scale_stats)
            for g in game_log_dicts:
                gamelog_rows.append({
                    "選手名": name,
                    "日付": g["date"], "対戦": g["opponent"], "結果": g["result"],
                    "空振り率": g["swstr_pct"], "ボール球SW%": g["chase_pct"],
                    "ストライク率": g["strike_pct"], "ゾーン率": g["zone_pct"],
                    "K-BB%": g["kbb_pct"], "K%": g["k_pct"], "BB%": g["bb_pct"], "ゴロ率": g["gb_pct"],
                    "投球回": g["ip"], "球数": g["pitches"], "奪三振": g["k"], "与四球": g["bb"],
                    "被安打": g["h"], "自責点": g["er"],
                    "pitch_detail_json": json.dumps(g["pitch_detail"], ensure_ascii=False),
                    "course_detail_json": json.dumps(g["course_detail"], ensure_ascii=False),
                })

            if numeric_json_dir:
                pitcher_hand = appearances[0]["player"].get("hand") if appearances else None
                numeric_cards[name] = {
                    "name": name,
                    "team": team,
                    "role": season["役割"],
                    "throws": pitcher_hand,  # 投手の利き腕 R/L（配球プロットの表示用フリップに使う）
                    "games": season["登板数"],
                    "innings": season["投球回"],
                    "era": season["防御率"],
                    "k_bb_pct": season["K-BB%"],
                    "gb_pct": season["ゴロ率"],
                    "k": season["奪三振"],
                    "bb": season["与四球"],
                    "swstr_pct_season": season["空振り率"],
                    "chase_pct_season": season["ゾーン外スイング率"],
                    "strike_pct_season": season["ストライク率"],
                    "zone_pct_season": season["ゾーン率"],
                    "kbb_pct_season": season["K-BB%"],
                    "k_pct_season": season["K%"],
                    "bb_pct_season": season["BB%"],
                    "pitch_evaluations_numeric": pitch_numeric_rows,
                    "season_pitch_detail": season_pitch_detail,
                    # 球種比較テーブルの「合計」行用の、本当のシーズン実数（球種別内訳からの
                    # 再加重平均ではない）。all は season（calc_season_stats）と同じ値。
                    "season_totals": {
                        "all": {
                            "空振り率": season["空振り率"],
                            "ゾーン外スイング率": season["ゾーン外スイング率"],
                            "ストライク率": season["ストライク率"],
                            "ゾーン率": season["ゾーン率"],
                            "ゴロ率": season["ゴロ率"],
                        },
                        "vsR": season_by_hand["vsR"],
                        "vsL": season_by_hand["vsL"],
                    },
                    "season_course_detail": season_course_detail,
                    "season_course_locs": season_course_locs,
                    "game_log": game_log_dicts,
                    # rankingsはこの後、全選手分揃ってから付与する
                }
        except Exception as e:
            print(f"  [SKIP] {name}: 集計中にエラーのためスキップ({type(e).__name__}: {e})")
            continue

    # 順位（防御率・K-BB%・K%・BB%・ゴロ率）を算出し、シーズン集計に列として付与
    rankings = compute_rankings(season_rows, rank_min_ip=RANK_MIN_IP)
    for row in season_rows:
        rk = rankings.get(row["選手名"], {})
        row["防御率_順位"] = rk.get("era", {}).get("rank")
        row["防御率_順位_母数"] = rk.get("era", {}).get("total")
        row["K-BB%_順位"] = rk.get("k_bb_pct", {}).get("rank")
        row["K-BB%_順位_母数"] = rk.get("k_bb_pct", {}).get("total")
        row["K%_順位"] = rk.get("k_pct", {}).get("rank")
        row["K%_順位_母数"] = rk.get("k_pct", {}).get("total")
        row["BB%_順位"] = rk.get("bb_pct", {}).get("rank")
        row["BB%_順位_母数"] = rk.get("bb_pct", {}).get("total")
        row["ゴロ率_順位"] = rk.get("gb_pct", {}).get("rank")
        row["ゴロ率_順位_母数"] = rk.get("gb_pct", {}).get("total")

    # 球種別詳細シート向け：全体/対右/対左の球種別順位を算出し、mix_rowsに列として付与
    role_map = {row["選手名"]: row.get("役割") for row in season_rows}
    ip_map = {row["選手名"]: row.get("投球回") for row in season_rows}
    compute_pitch_rankings(mix_rows, role_map, ip_map)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame(season_rows).to_excel(writer, sheet_name="シーズン集計", index=False)
        pd.DataFrame(mix_rows).to_excel(writer, sheet_name="球種別詳細", index=False)
        pd.DataFrame(course_rows).to_excel(writer, sheet_name="コース分布", index=False)
        pd.DataFrame(course_grid_rows).to_excel(writer, sheet_name="コースグリッド25分割", index=False)
        pd.DataFrame(count_rows).to_excel(writer, sheet_name="カウント別パターン", index=False)
        pd.DataFrame(gamelog_rows).to_excel(writer, sheet_name="試合ログ", index=False)

    if numeric_json_dir:
        os.makedirs(numeric_json_dir, exist_ok=True)
        index_players = []
        for name, card in numeric_cards.items():
            card["rankings"] = rankings.get(name, {})
            card["categories"] = classify_pitcher_categories(card)
            player_id = _slugify_name(name)
            with open(os.path.join(numeric_json_dir, f"{player_id}.json"), "w", encoding="utf-8") as f:
                json.dump(card, f, ensure_ascii=False)
            ip_float = _ip_to_outs(card.get("innings")) / 3
            index_players.append({
                "id": player_id, "name": name, "team": card.get("team"), "role": card.get("role"),
                "innings": card.get("innings"), "innings_num": round(ip_float, 1),
                "categories": card["categories"],
                # 選手一覧での並び替え用に主要指標もindex.jsonに含めておく
                # （選手ごとのJSONを全件取得しなくても一覧画面でソートできるようにするため）
                "era": card.get("era"),
                "k_bb_pct": card.get("kbb_pct_season"),
                "k_pct": card.get("k_pct_season"),
                "bb_pct": card.get("bb_pct_season"),
                "swstr_pct": card.get("swstr_pct_season"),
                "gb_pct": card.get("gb_pct"),
            })
        with open(os.path.join(numeric_json_dir, "index.json"), "w", encoding="utf-8") as f:
            json.dump({"players": index_players}, f, ensure_ascii=False, indent=2)
        print(f"  数値JSON: {len(index_players)}選手分を {numeric_json_dir} に出力")

    return out_path


def _slugify_name(name: str) -> str:
    """選手名からファイル名用のIDを作る（英数字以外はアンダースコアに置換）"""
    s = re.sub(r"[^\w]+", "_", name.strip().lower())
    return s.strip("_") or "unknown"


# ==================================================
# Section 7. CLI
# ==================================================

def parse_args():
    p = argparse.ArgumentParser(description="LLM入力用xlsx（シーズン投手データ）を生成する")
    p.add_argument("--games-json-dir", required=True, help="games/json/{date}.json が入っているディレクトリ")
    p.add_argument("--out", required=True, help="出力xlsxのパス")
    p.add_argument("--min-ip", type=float, default=0.0, help="この投球回以上の投手のみ対象にする（デフォルト0=全員）")
    p.add_argument("--players", nargs="*", default=None, help="対象選手名を絞り込む場合はスペース区切りで指定（省略時は全投手）")
    p.add_argument("--numeric-json-dir", default=None,
                    help="指定すると、選手ごとの数値データJSON（pitcher_cards_numeric/配下）とindex.jsonも出力する")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    out = export_llm_input_xlsx(
        games_json_dir=args.games_json_dir,
        out_path=args.out,
        min_ip=args.min_ip,
        target_names=args.players,
        numeric_json_dir=args.numeric_json_dir,
    )
    print(f"✅ 出力完了: {out}")
