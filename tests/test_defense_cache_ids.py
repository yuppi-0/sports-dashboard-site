import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import export_llm_input_batter as eb  # noqa: E402


def test_same_name_players_keep_their_own_running_and_war(tmp_path):
    p = tmp_path / "def.xlsx"
    with pd.ExcelWriter(p) as w:
        pd.DataFrame([{"name": "smith, will", "player_id": 669257, "sb": 3, "cs": 1},
                      {"name": "smith, will", "player_id": 519293, "sb": 0, "cs": 0},
                      {"name": "judge, aaron", "player_id": 592450, "sb": 4, "cs": 0}]).to_excel(w, sheet_name="Running", index=False)
        pd.DataFrame([{"name": "smith, will", "player_id": 669257, "war": 3.3},
                      {"name": "smith, will", "player_id": 519293, "war": 0.1}]).to_excel(w, sheet_name="WAR", index=False)
    cache = eb.load_mlb_defense_cache(str(p))
    assert eb.defense_lookup(cache, "will smith", 669257)["sb"] == 3
    assert eb.defense_lookup(cache, "will smith", 669257)["war"] == 3.3
    assert eb.defense_lookup(cache, "will smith", 519293)["sb"] == 0
    assert eb.defense_lookup(cache, "aaron judge", 592450)["sb"] == 4     # 同名衝突の無い選手は従来どおり名前で引ける
    assert eb.defense_lookup(cache, "aaron judge")["sb"] == 4
