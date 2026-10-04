"""保存済み箱スコアの訂正漏れチェックの比較ロジックのテスト（通信なし）。"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import check_npb_stale_games as chk  # noqa: E402


def df(rows):
    return pd.DataFrame(rows)


def test_numeric_format_difference_is_not_a_change():
    old = df([{"チーム": "阪神", "選手名": "村上", "投球回": 6.0, "被安打": 3}])
    new = df([{"チーム": "阪神", "選手名": "村上", "投球回": "6", "被安打": "3"}])
    assert chk.diff_sheet(old, new, ["投球回", "被安打"], False) == []


def test_real_change_is_reported_including_inning_cells():
    old = df([{"チーム": "阪神", "選手名": "佐藤", "安打": 2, "1回": "左２", "3回": "遊ゴ"}])
    new = df([{"チーム": "阪神", "選手名": "佐藤", "安打": 3, "1回": "左２", "3回": "左安"}])
    out = chk.diff_sheet(old, new, ["安打"], True)
    assert {(c, a, b) for _k, c, a, b in out} == {("安打", "2", "3"), ("3回", "遊ゴ", "左安")}
