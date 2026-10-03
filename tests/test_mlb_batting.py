"""MLB打者の打点・打数まわりのテスト（通信なし）。"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run_mlb as mlb  # noqa: E402


def last_rows(rows):
    return pd.DataFrame(rows)


def test_rbi_follows_official_rules():
    df = last_rows([
        {"events": "home_run", "bat_score": 0, "post_bat_score": 2, "outs_when_up": 0, "on_3b": 5},               # 本塁打2点 → 2
        {"events": "grounded_into_double_play", "bat_score": 1, "post_bat_score": 2, "outs_when_up": 0, "on_3b": 5},  # 併殺打で入った得点 → 0
        {"events": "field_error", "bat_score": 0, "post_bat_score": 1, "outs_when_up": 1, "on_3b": 7},           # 失策だが一死三塁 → 1
        {"events": "field_error", "bat_score": 0, "post_bat_score": 1, "outs_when_up": 0, "on_3b": np.nan},       # 失策・三塁走者なし → 0
        {"events": "walk", "bat_score": 3, "post_bat_score": 4, "outs_when_up": 1, "on_3b": 9},                   # 押し出し → 1
        {"events": "sac_fly", "bat_score": 2, "post_bat_score": 3, "outs_when_up": 1, "on_3b": 9},                # 犠飛 → 1
    ])
    assert list(mlb._pa_rbi(df)) == [2, 0, 1, 0, 1, 1]


def test_rbi_uses_statsapi_value_as_is():
    df = last_rows([{"events": "grounded_into_double_play", "bat_score": 0, "post_bat_score": 1, "_statsapi_source": True}])
    assert list(mlb._pa_rbi(df)) == [1]


def test_rbi_without_score_columns_is_zero():
    assert list(mlb._pa_rbi(pd.DataFrame({"events": ["single"]}))) == [0]
