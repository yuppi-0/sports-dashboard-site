"""同姓同名の別人を別カードに分けるテスト（通信なし）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import player_identity as pi  # noqa: E402


def game(home, away, home_names, away_names, key="batters"):
    return {"home": home, "away": away, key: {"home": [{"name": n} for n in home_names], "away": [{"name": n} for n in away_names]}}


def test_same_name_on_different_teams_same_day_is_split():
    data = {"2025-06-01": [game("LAD", "ATH", ["Max Muncy"], ["Max Muncy"])],
            "2025-06-02": [game("LAD", "SEA", ["Max Muncy"], ["Ty France"])]}
    out = pi.disambiguate_same_name(data, "batters")
    assert out == {"max muncy": ["ATH", "LAD"]}
    assert data["2025-06-01"][0]["batters"]["home"][0]["name"] == "Max Muncy (LAD)"
    assert data["2025-06-01"][0]["batters"]["away"][0]["name"] == "Max Muncy (ATH)"
    assert data["2025-06-02"][0]["batters"]["home"][0]["name"] == "Max Muncy (LAD)"
    assert data["2025-06-02"][0]["batters"]["away"][0]["name"] == "Ty France"


def test_traded_player_stays_one_card():
    data = {"2025-06-01": [game("LAD", "SEA", ["Jose Iglesias"], [])],
            "2025-07-01": [game("NYM", "SEA", ["Jose Iglesias"], [])]}
    assert pi.disambiguate_same_name(data, "batters") == {}
    assert data["2025-07-01"][0]["batters"]["home"][0]["name"] == "Jose Iglesias"


def test_accent_variants_count_as_same_name():
    data = {"2025-06-01": [game("HOU", "WSH", ["Luis Garcia"], ["Luis García"], key="pitchers")]}
    assert list(pi.disambiguate_same_name(data, "pitchers")) == ["luis garcia"]
