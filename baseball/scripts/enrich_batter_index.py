"""既存の打者 index.json に、順位の付け直し用の追加指標（index_fields.BATTER_INDEX_EXTRA）を足す（docs/ 側だけを書き換える）。

  python baseball/scripts/enrich_batter_index.py [--docs docs]

各 batter_cards_numeric/ の選手カードJSONを読み、その年度の overall から追加指標を index.json の各選手に入れる。
以降の自動更新は export_llm_input_batter.py が同じ指標を書くので、この作業は既存データに対する1回きりでよい。
"""
from __future__ import annotations
import argparse
import gzip
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from index_fields import batter_index_extra  # noqa: E402


def load_card(d: Path, pid: str):
    for name, opener in ((f"{pid}.json.gz", lambda p: gzip.open(p, "rt", encoding="utf-8")), (f"{pid}.json", lambda p: open(p, encoding="utf-8"))):
        f = d / name
        if f.exists():
            with opener(f) as fh:
                return json.load(fh)
    return None


def enrich_dir(d: Path) -> tuple:
    idx_path = d / "index.json"
    data = json.loads(idx_path.read_text(encoding="utf-8"))
    n = 0
    for p in data.get("players", []):
        card = load_card(d, p["id"])
        if not card:
            continue
        seasons = card.get("seasons") or {}
        year = card.get("latestYear") or (max(seasons) if seasons else None)
        extra = batter_index_extra((seasons.get(year) or {}).get("overall"))
        if extra:
            p.update(extra)
            n += 1
    idx_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(data.get("players", [])), n


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docs", default="docs")
    a = ap.parse_args()
    dirs = sorted(Path(a.docs).glob("baseball/data/*/*年/*/*/batter_cards_numeric")) + sorted(Path(a.docs).glob("baseball/data/*/*年/*/batter_cards_numeric"))   # NPB（1軍/2軍）とMLB
    for d in dirs:
        total, n = enrich_dir(d)
        print(f"{d.relative_to(a.docs)}: {total}人中 {n}人に追加指標を入れた")


if __name__ == "__main__":
    main()
