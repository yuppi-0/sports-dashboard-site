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


def test_official_batting_replaces_recomputed_line():
    stats = {"打点": 3, "打席": 5, "打数": 4, "安打": 2, "二塁打": 0, "三塁打": 0, "本塁打": 0, "四球": 1, "死球": 0, "三振": 1,
             "犠飛": 0, "犠打": 0, "盗塁": 1, "打球": "x"}
    box = {"plateAppearances": 5, "atBats": 3, "hits": 2, "doubles": 1, "triples": 0, "homeRuns": 0,
           "baseOnBalls": 1, "hitByPitch": 0, "strikeOuts": 1, "sacFlies": 1, "sacBunts": 0, "rbi": 2}
    out = mlb.apply_official_batting(stats, box)
    assert (out["打席"], out["打数"], out["安打"], out["二塁打"], out["犠飛"], out["打点"]) == (5, 3, 2, 1, 1, 2)
    assert out["単打"] == 1 and out["長打"] == 1
    assert out["打率"] == round(2 / 3, 3) and out["長打率"] == round(3 / 3, 3)
    assert out["出塁率"] == round((2 + 1) / (3 + 1 + 0 + 1), 3)
    assert out["盗塁"] == 1 and out["打球"] == "x"        # 箱スコアに無い項目はそのまま
    assert stats["打点"] == 3                              # 元の辞書は変更しない


def test_official_batting_falls_back():
    stats = {"打点": 3, "打席": 5}
    assert mlb.apply_official_batting(stats, None) is stats
    assert mlb.apply_official_batting(stats, {}) is stats
    assert mlb.apply_official_batting(stats, {"rbi": 2})["打点"] == 2     # 必須項目が無ければ打点だけ
    assert mlb.apply_official_batting(stats, {"rbi": None}) is stats


def test_official_zero_batters_faced_is_respected():
    # 打席の途中で降板した先発は公式の対戦打者数が0（その打席は救援投手に付く）。再集計値（1）を残さない
    box = {"inningsPitched": "0.0", "runs": 0, "earnedRuns": 0, "hits": 0, "homeRuns": 0,
           "baseOnBalls": 0, "hitByPitch": 0, "strikeOuts": 0, "battersFaced": 0}
    out = mlb.apply_official_pitching({**RECOMPUTED, "対戦打者数": 1}, box)
    assert out["対戦打者数"] == 0
