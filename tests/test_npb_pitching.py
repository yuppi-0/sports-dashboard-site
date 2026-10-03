"""NPB投手の試合別 対戦打者数・K%・BB% が公式の箱スコアに基づくことのテスト（通信なし）。"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run as npb  # noqa: E402


def test_uses_box_batters_faced_instead_of_pitch_derived():
    # 公式: 13人と対戦し2奪三振・0四球。投球データ再集計は12人（1打席取りこぼし）
    row = pd.Series({"打者": 13, "奪三振": 2, "与四球": 0})
    tbf, k, bb, kbb = npb.official_pitcher_rates(row, 12, 16.7, 0.0, 16.7)
    assert (tbf, k, bb, kbb) == (13, 15.4, 0.0, 15.4)


def test_falls_back_to_pitch_derived_when_box_unreadable():
    for row in (pd.Series({"奪三振": 2, "与四球": 0}), pd.Series({"打者": "", "奪三振": 2, "与四球": 0}),
                pd.Series({"打者": 0, "奪三振": 2, "与四球": 0})):
        assert npb.official_pitcher_rates(row, 12, 16.7, 0.0, 16.7) == (12, 16.7, 0.0, 16.7)


def test_kbb_uses_walks():
    row = pd.Series({"打者": 10, "奪三振": 3, "与四球": 1})
    assert npb.official_pitcher_rates(row, 9, 0, 0, 0) == (10, 30.0, 10.0, 20.0)
