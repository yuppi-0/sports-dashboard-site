"""アクセント有無の表記違い（速報由来 / Statcast由来）が同一選手として扱われるテスト（通信なし）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import export_llm_input as ep  # noqa: E402
import export_llm_input_batter as eb  # noqa: E402


def _data(side_key, names):
    return {"2026-05-01": [{side_key: {"home": [{"name": n} for n in names], "away": []}}]}


def test_pitcher_accent_variants_are_one_player():
    data = _data("pitchers", ["Jose A. Ferrer", "José A. Ferrer", "Ferrer, José A."])
    assert ep.build_all_pitcher_names(data) == {"José A. Ferrer"}
    assert len(ep.build_appearances(data, "José A. Ferrer")) == 3
    assert len(ep.build_appearances(data, "Jose A. Ferrer")) == 3


def test_batter_accent_variants_are_one_player():
    data = _data("batters", ["Yainer Diaz", "Yainer Díaz"])
    assert eb.build_all_batter_names(data) == {"Yainer Díaz"}
    assert len(eb.build_appearances_batter(data, "Yainer Diaz")) == 2


def test_different_players_stay_separate():
    data = _data("pitchers", ["Luis Garcia", "Luis Gil"])
    assert ep.build_all_pitcher_names(data) == {"Luis Garcia", "Luis Gil"}


def test_pitching_only_games_are_counted_for_two_way_players():
    data = {
        "2026-06-01": [{"home": "日本ハム", "away": "楽天",
                        "batters": {"home": [{"name": "柴田 獅子"}], "away": []},
                        "pitchers": {"home": [{"name": "柴田 獅子"}], "away": []}}],          # 打席あり＝数えない
        "2026-06-02": [{"home": "日本ハム", "away": "楽天",
                        "batters": {"home": [{"name": "別の 選手"}], "away": []},
                        "pitchers": {"home": [{"name": "柴田 獅子"}], "away": []}}],          # 投手だけ＝数える
    }
    assert eb.count_pitching_only_games(data, "柴田 獅子") == 1
    assert eb.count_pitching_only_games(data, "別の 選手") == 0
