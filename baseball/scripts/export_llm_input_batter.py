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
  - 日別JSON側の打者エントリ（run.py の build_batter()）は「1試合1行」の集計値
    しか持っておらず、投手版のような球種別・投球コース別の内訳は無い。
    そのため球種別・コース別の分析は今回のスコープに含めていない。
  - 対右投手/対左投手のスプリットは、打席ごとの対戦投手までは紐付いていないため、
    「その試合の相手チーム先発投手の投球腕」をその試合全体の対戦相手の腕として
    近似的に分類している（登板交代で左右が変わった分は反映できない）。
  - 出塁率/長打率/OPSは、単打・二塁打・三塁打の内訳が日別JSON側に無く総塁打数を
    正確に積み上げられないため、試合ごとに計算済みの値を打席数で加重平均した近似値。
    安打・本塁打・四死球・三振・打点・盗塁・打率・K%・BB%は実数の積み上げなので正確。
  - 守備・走塁の高度指標（守備率・レンジファクター・盗塁死・盗塁成功率など）は、
    現時点のrun.pyに守備成績ページのスクレイピングが実装されていないため取得できない。
    numeric_json側にはキーだけ用意してNoneを入れてある。追加するには、run.py側に
    守備成績（刺殺・補殺・失策・捕逸）と盗塁死を取得する新しいスクレイピングの
    ステップを追加する必要がある（別途対応）。

使い方:
    python export_llm_input_batter.py \
        --games-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/games/json" \
        --out "data/datamart/llm_input/プロ野球_2026_打者データ.xlsx" \
        --numeric-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/batter_cards_numeric" \
        --min-pa 0
"""

from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter

import pandas as pd

# 日別JSONの読み込みロジックは投手版と完全に共通のため使い回す
from export_llm_input import load_daily_games


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
    run.py側で pitchers.home/away は (役割!=先発, 投手試合ID) 順にソート済みのため、
    先頭要素=その試合の先発投手とみなせる。取れなければNone。"""
    pitchers = game.get("pitchers") or {}
    if not isinstance(pitchers, dict):
        return None
    opp_side = "away" if batter_side == "home" else "home"
    opp_list = pitchers.get(opp_side) or []
    if not opp_list or not isinstance(opp_list[0], dict):
        return None
    return opp_list[0].get("hand")


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
            "三振": 0, "打点": 0, "盗塁": 0, "打率": None, "出塁率": None, "長打率": None,
            "OPS": None, "K%": None, "BB%": None,
            "O-Swing%": None, "Z-Swing%": None, "whiff%": None,
        }

    def s(key):
        return sum((ap["player"].get(key) or 0) for ap in rows)

    pa, ab, h, hr = s("pa"), s("ab"), s("h"), s("hr")
    bb, k, rbi, sb = s("bb"), s("k"), s("rbi"), s("sb")

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
        "四死球": bb, "三振": k, "打点": rbi, "盗塁": sb,
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
            "order": p.get("order"), "pos": p.get("pos"),
            "pa": p.get("pa"), "ab": p.get("ab"), "h": p.get("h"), "hr": p.get("hr"),
            "bb": p.get("bb"), "k": p.get("k"), "rbi": p.get("rbi"), "sb": p.get("sb"),
            "avg_game": round((p.get("h") or 0) / ab, 3) if ab > 0 else None,
            "obp": p.get("obp"), "slg": p.get("slg"), "ops": p.get("ops"),
            "abs": p.get("abs"),
        })
    return rows


# ==================================================
# Section 4. 順位算出・カテゴリタグ
# ==================================================

RANK_MIN_PA = 100  # 順位算出の資格打席（この打席数未満の選手は順位母集団から除外）

# (シーズン集計側の列名, numeric json側のキー名, 高いほど良いか)
_RANK_SPECS = [
    ("打率", "avg", True), ("出塁率", "obp", True), ("長打率", "slg", True), ("OPS", "ops", True),
    ("本塁打", "hr", True), ("盗塁", "sb", True), ("K%", "k_pct", False), ("BB%", "bb_pct", True),
]


def compute_batter_rankings(season_rows: list[dict], rank_min_pa: float = RANK_MIN_PA) -> dict:
    """打率・出塁率・長打率・OPS・本塁打・盗塁・K%・BB%の順位を算出する。
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


# ==================================================
# Section 5. メイン: xlsx / numeric json 出力
# ==================================================

def export_llm_input_batter_xlsx(games_json_dir: str, out_path: str, min_pa: float = 0.0,
                                  target_names: list[str] | None = None,
                                  numeric_json_dir: str | None = None) -> str:
    """
    numeric_json_dir を指定すると、xlsxに加えて選手ごとの数値データJSON
    （batter_cards_numeric/{選手ID}.json）と選手一覧 index.json も書き出す。
    これらはbatter-cards.html側がそのまま読み込む「数値だけ」のデータ。
    """
    all_data = load_daily_games(games_json_dir)
    names = set(target_names) if target_names else build_all_batter_names(all_data)

    season_rows, split_rows, gamelog_rows = [], [], []
    numeric_cards: dict[str, dict] = {}

    for name in sorted(names):
        try:
            appearances = build_appearances_batter(all_data, name)
            if not appearances:
                continue

            season = calc_season_batter_stats(appearances)
            season["選手名"] = name

            if (season.get("打席") or 0) < min_pa:
                continue

            vs_r = calc_season_batter_stats(appearances, hand_filter="R")
            vs_l = calc_season_batter_stats(appearances, hand_filter="L")

            pos = determine_primary_position(appearances)
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
                numeric_cards[name] = {
                    "name": name, "team": team, "pos": pos,
                    "games": season["試合数"], "pa": season["打席"], "ab": season["打数"],
                    "h": season["安打"], "hr": season["本塁打"], "bb": season["四死球"],
                    "k": season["三振"], "rbi": season["打点"], "sb": season["盗塁"],
                    "avg": season["打率"], "obp": season["出塁率"], "slg": season["長打率"],
                    "ops": season["OPS"],
                    "k_pct_season": season["K%"], "bb_pct_season": season["BB%"],
                    "chase_pct_season": season["O-Swing%"],
                    "contact_pct_season": season["Z-Swing%"],
                    "whiff_pct_season": season["whiff%"],
                    "splits": {"vsR": vs_r, "vsL": vs_l},
                    "game_log": game_log_dicts,
                    # 守備・走塁の高度指標: 現状のrun.pyには守備成績スクレイピングが
                    # 実装されていないため取得不可。将来対応するまではNoneのまま出力する。
                    "defense": {
                        "fielding_pct": None,  # 守備率（要: 守備成績ページの新規スクレイピング）
                        "range_factor": None,  # レンジファクター（同上）
                        "cs": None,             # 盗塁死（同上）
                        "sb_success_pct": None,  # 盗塁成功率（盗塁死が無いため算出不可）
                    },
                    # rankings/categoriesはこの後、全選手分揃ってから付与する
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
        for name, card in numeric_cards.items():
            card["rankings"] = rankings.get(name, {})
            card["categories"] = classify_batter_categories(card)
            player_id = _slugify_name(name)
            with open(os.path.join(numeric_json_dir, f"{player_id}.json"), "w", encoding="utf-8") as f:
                json.dump(card, f, ensure_ascii=False)
            index_players.append({
                "id": player_id, "name": name, "team": card.get("team"), "pos": card.get("pos"),
                "categories": card["categories"],
                "games": card.get("games"), "pa": card.get("pa"),
                "avg": card.get("avg"), "obp": card.get("obp"), "slg": card.get("slg"),
                "ops": card.get("ops"),
                "hr": card.get("hr"), "rbi": card.get("rbi"), "sb": card.get("sb"),
                "k_pct": card.get("k_pct_season"), "bb_pct": card.get("bb_pct_season"),
            })
        with open(os.path.join(numeric_json_dir, "index.json"), "w", encoding="utf-8") as f:
            json.dump({"players": index_players}, f, ensure_ascii=False, indent=2)
        print(f"  数値JSON: {len(index_players)}選手分を {numeric_json_dir} に出力")

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
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    out = export_llm_input_batter_xlsx(
        games_json_dir=args.games_json_dir,
        out_path=args.out,
        min_pa=args.min_pa,
        target_names=args.players,
        numeric_json_dir=args.numeric_json_dir,
    )
    print(f"✅ 出力完了: {out}")
