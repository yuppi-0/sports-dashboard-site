"""選手一覧に載らない古いカードJSONの掃除のテスト（通信なし）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import jsonio  # noqa: E402


def test_prune_removes_only_orphans(tmp_path):
    for name in ("jose_a_ferrer.json.gz", "josé_a_ferrer.json.gz", "657675.json.gz", "index.json", "pivot_population_2026.json.gz"):
        (tmp_path / name).write_bytes(b"x")
    removed = jsonio.prune_orphan_cards(str(tmp_path), {"josé_a_ferrer"})
    assert removed == ["657675", "jose_a_ferrer"]
    left = sorted(p.name for p in tmp_path.iterdir())
    assert left == ["index.json", "josé_a_ferrer.json.gz", "pivot_population_2026.json.gz"]
