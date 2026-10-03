"""MLB投手の試合別成績を公式の箱スコアで上書きするロジックのテスト（通信なし）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run_mlb as mlb  # noqa: E402


RECOMPUTED = {"投球回": "1.0", "投球数": 20, "対戦打者数": 5, "失点": 3, "自責点": 3, "被安打": 2, "被本塁打": 0,
              "与四球": 1, "与死球": 0, "奪三振": 1, "K%": 20.0, "BB%": 20.0, "K-BB%": 0.0, "GB": 2}


def test_official_box_overrides_recomputed_line():
    # 継投で引き継いだ走者の失点・非自責の失点は公式の箱スコアでは別に数える
    box = {"inningsPitched": "1.1", "runs": 3, "earnedRuns": 1, "hits": 2, "homeRuns": 0,
           "baseOnBalls": 1, "hitBatsmen": 0, "strikeOuts": 1, "battersFaced": 6}
    out = mlb.apply_official_pitching(RECOMPUTED, box)
    assert (out["投球回"], out["失点"], out["自責点"]) == ("1.1", 3, 1)
    assert out["対戦打者数"] == 6
    assert (out["K%"], out["BB%"], out["K-BB%"]) == (16.7, 16.7, 0.0)
    assert out["GB"] == 2                      # 打球・スイング指標は投球データのまま
    assert RECOMPUTED["自責点"] == 3           # 元の辞書は変更しない


def test_falls_back_when_box_missing_or_unreadable():
    assert mlb.apply_official_pitching(RECOMPUTED, None) is RECOMPUTED
    assert mlb.apply_official_pitching(RECOMPUTED, {}) is RECOMPUTED
    assert mlb.apply_official_pitching(RECOMPUTED, {"inningsPitched": ""}) is RECOMPUTED


def test_official_rbi_replaces_rule_based_value():
    stats = {"打点": 3, "安打": 2}
    assert mlb.apply_official_rbi(stats, {"rbi": 2})["打点"] == 2
    assert stats["打点"] == 3                               # 元の辞書は変更しない
    assert mlb.apply_official_rbi(stats, None) is stats
    assert mlb.apply_official_rbi(stats, {"rbi": None}) is stats
    assert mlb.apply_official_rbi(stats, {"rbi": "x"}) is stats
