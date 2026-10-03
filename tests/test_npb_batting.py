"""NPB打者の打率・出塁率・長打率・OPSが公式の値になることのテスト（通信なし）。"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run as npb  # noqa: E402
import export_llm_input_batter as eb  # noqa: E402


def box(**kw):
    base = {"打数": 4, "得点": 1, "安打": 2, "打点": 1, "三振": 1, "四球": 0, "死球": 0, "犠打": 0, "盗塁": 0, "失策": 0, "本塁打": 0}
    base.update(kw)
    return pd.Series(base)


def test_box_counts_read_official_columns_and_inning_cells():
    row = box(**{"打数": 3, "安打": 2, "四球": 1, "本塁打": 0, "1回": "左２", "3回": "中３", "5回": "右犠飛", "7回": "空三振", "8回": "四球"})
    c = npb.box_batting_counts(row)
    assert (c["ab"], c["h"], c["bb"], c["d2"], c["d3"], c["sf"]) == (3, 2, 1, 1, 1, 1)
    assert c["pa"] == 3 + 1 + 0 + 0 + 1      # 打数+四球+死球+犠打+犠飛


def test_box_counts_returns_none_when_columns_missing():
    assert npb.box_batting_counts(pd.Series({"選手名": "x"})) is None


def test_season_rates_use_official_definitions():
    # 山口航輝（2026-10-02時点の公式成績）：打数396 安打109 二塁打18 三塁打1 本塁打34 四球26 死球6 犠飛2
    one = {"pa": 430, "ab": 396, "h": 109, "hr": 34, "bb": 32, "k": 100, "rbi": 76, "sb": 0,
           "hbp": 6, "bbh": 32, "dbl": 18, "tpl": 1, "sf": 2, "sh": 0}
    r = eb.calc_season_batter_stats([{"player": one}])
    assert (r["打率"], r["出塁率"], r["長打率"], r["OPS"]) == (0.275, 0.328, 0.583, 0.911)


def test_season_rates_sum_over_games_not_average_of_rates():
    g1 = {"pa": 5, "ab": 4, "h": 4, "hr": 0, "bb": 1, "k": 0, "rbi": 0, "sb": 0, "hbp": 0, "bbh": 1, "dbl": 0, "tpl": 0, "sf": 0, "sh": 0}
    g2 = {"pa": 5, "ab": 5, "h": 0, "hr": 0, "bb": 0, "k": 0, "rbi": 0, "sb": 0, "hbp": 0, "bbh": 0, "dbl": 0, "tpl": 0, "sf": 0, "sh": 0}
    r = eb.calc_season_batter_stats([{"player": g1}, {"player": g2}])
    assert r["打率"] == round(4 / 9, 3)
    assert r["出塁率"] == round(5 / 10, 3)
    assert r["長打率"] == round(4 / 9, 3)


def test_falls_back_to_weighted_average_when_breakdown_missing():
    g = {"pa": 4, "ab": 4, "h": 1, "hr": 0, "bb": 0, "k": 0, "rbi": 0, "sb": 0, "obp": 0.25, "slg": 0.25, "ops": 0.5}
    r = eb.calc_season_batter_stats([{"player": g}])
    assert (r["出塁率"], r["長打率"], r["OPS"]) == (0.25, 0.25, 0.5)


def test_ops_sums_unrounded_obp_and_slg():
    # 佐藤輝明: 出塁率.391 + 長打率.625 の丸め値の和は1.016だが、公式のOPSは丸める前の合算で1.017
    one = {"pa": 598, "ab": 523, "h": 164, "hr": 40, "bb": 66, "k": 100, "rbi": 105, "sb": 0,
           "hbp": 6, "bbh": 69, "dbl": 22, "tpl": 1, "sf": 0, "sh": 0}
    r = eb.calc_season_batter_stats([{"player": one}])
    exact = (164 + 69) / (523 + 69) + (164 + 22 + 2 + 120) / 523
    assert r["OPS"] == round(exact, 3)
