import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run_mlb as rm  # noqa: E402


def test_player_on_both_sides_of_suspended_game_is_summed():
    a = {"plateAppearances": 0, "atBats": 0, "hits": 0, "summary": "0-0"}
    b = {"plateAppearances": 4, "atBats": 4, "hits": 1, "summary": "1-4"}
    m = rm._merge_box_stats(a, b)
    assert m["plateAppearances"] == 4 and m["hits"] == 1
    assert rm._merge_box_stats(b, a)["plateAppearances"] == 4          # 順序に依らない
    assert rm._merge_box_stats({"inningsPitched": "0.2"}, {"inningsPitched": "0.2"})["inningsPitched"] == "1.1"


def test_pitcher_without_any_pitch_is_added_from_boxscore(monkeypatch):
    box = {"pitching": {1: {"inningsPitched": "1.0", "battersFaced": 4, "strikeOuts": 1, "hits": 1},
                        2: {"inningsPitched": "0.1", "battersFaced": 0, "strikeOuts": 0},
                        3: {"inningsPitched": "0.0", "battersFaced": 0}},
           "batting": {}, "side": {1: "away", 2: "away", 3: "home"}, "name": {1: "Ann Lee", 2: "Joe Smith", 3: "Bob Roe"}}
    monkeypatch.setattr(rm, "fetch_boxscore", lambda gid: box)
    df = pd.DataFrame({"game_pk": [99], "home_team": ["BOS"], "away_team": ["SEA"]})
    rows = [{"試合ID": "99", "試合日": "2021-07-26", "選手名": "lee, ann", "投手ID": 1, "チーム": "SEA", "ホーム/アウェイ": "away",
             "役割": "先発", "勝敗成績": "", "投球回": "1.0", "投球数": 20, "被安打": 1, "K%": 25.0}]
    extra = rm._boxscore_only_pitcher_rows(df, rows)
    assert len(extra) == 1                                              # 投球データのあるID=1と、登板していないID=3は足さない
    r = extra[0]
    assert r["投手ID"] == 2 and r["選手名"] == "smith, joe" and r["チーム"] == "SEA" and r["投球回"] == "0.1"
    assert r["対戦打者数"] == 0 and r["投球数"] == 0
