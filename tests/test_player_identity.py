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


def gid(home, away, home_players, away_players, key="pitchers"):
    return {"home": home, "away": away, key: {"home": home_players, "away": away_players}}


def test_id_unifies_renamed_player_and_splits_same_name_different_ids():
    data = {
        "2026-05-21": [gid("NYM", "CIN", [{"name": "Thornton, Zach", "id": 1}], [{"name": "Muncy, Max", "id": 10}])],
        "2026-09-25": [gid("NYM", "ATH", [{"name": "Thornton, Zac", "id": 1}], [{"name": "Muncy, Max", "id": 11}])],
    }
    out = pi.canonicalize_by_id(data, "pitchers")
    assert out == {"max muncy": ["ATH", "CIN"]}
    assert data["2026-05-21"][0]["pitchers"]["home"][0]["name"] == "Zac Thornton"       # 最新の表記に統一
    assert data["2026-09-25"][0]["pitchers"]["home"][0]["name"] == "Zac Thornton"
    assert data["2026-05-21"][0]["pitchers"]["away"][0]["name"] == "Max Muncy (CIN)"
    assert data["2026-09-25"][0]["pitchers"]["away"][0]["name"] == "Max Muncy (ATH)"


def test_id_less_entries_are_left_alone():
    data = {"2026-05-21": [gid("NYM", "CIN", [{"name": "Thornton, Zach"}], [])]}
    assert pi.canonicalize_by_id(data, "pitchers") == {}
    assert data["2026-05-21"][0]["pitchers"]["home"][0]["name"] == "Thornton, Zach"


def test_pitcher_namesake_is_split_using_batter_side_conflict():
    """投手側だけでは同じ日に重ならない別人（野手登板が1試合だけのヤクルトのオスナ）も、打者側で同日に別チームと分かれば分ける。"""
    from player_identity import disambiguate_same_name, same_name_conflicts
    def game(date, home, away, batters, pitchers):
        return {"home": home, "away": away, "batters": {"home": [{"name": n} for n in batters.get("home", [])], "away": [{"name": n} for n in batters.get("away", [])]},
                "pitchers": {"home": [{"name": n} for n in pitchers.get("home", [])], "away": [{"name": n} for n in pitchers.get("away", [])]}}
    all_data = {
        "2026-05-01": [game("2026-05-01", "ヤクルト", "ソフトバンク", {"home": ["オスナ"], "away": ["オスナ"]}, {"away": ["オスナ"]})],
        "2026-05-02": [game("2026-05-02", "ヤクルト", "巨人", {}, {"home": ["オスナ"]})],      # 野手のオスナの登板（ヤクルト）
    }
    extra = same_name_conflicts(all_data, "batters")
    assert extra == {"オスナ": {"ヤクルト", "ソフトバンク"}}
    split = disambiguate_same_name(all_data, "pitchers", extra)
    assert "オスナ" in split
    names = {p["name"] for d in all_data.values() for g in d for side in ("home", "away") for p in g["pitchers"][side]}
    assert names == {"オスナ (ソフトバンク)", "オスナ (ヤクルト)"}
    assert disambiguate_same_name({"d": []}, "pitchers") == {}                                          # 余計な引数なしでも従来どおり
