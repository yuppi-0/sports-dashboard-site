"""同じ試合が別の日付にも保存されてしまっている「重複日付」を見つけて取り除く。

原因：試合の無い日（シーズン終了後など）は、Yahoo!の日程ページが「直近の試合日」の試合一覧を返すため、
その試合が翌日以降の日付としても保存され、日別データ→シーズン集計に二重・多重に入っていた
（取得処理は修正済み）。

判定：同じリーグ・同じ試合種別の中で、ある日付のRAW(all_games)の試合IDがすべて「それより前の日付」に
すでにある場合、その日付は重複とみなす（前の日付が本来の日）。一部だけ重なる日付は取り除かず報告だけ行う。

  python baseball/scripts/remove_duplicate_game_dates.py --year 2026            # 一覧のみ（書き換えない）
  python baseball/scripts/remove_duplicate_game_dates.py --year 2026 --apply    # RAW・日別データマート・日別JSONを削除
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pandas as pd

_SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_SCRIPT_DIR))
import run as npb  # noqa: E402

LEVELS = (("ichi", "1軍"), ("ni", "2軍"))


def game_ids_of(raw_date_dir: Path, date: str) -> set:
    p = raw_date_dir / f"all_games_{date}.xlsx"
    if not p.exists():
        return set()
    try:
        info = pd.read_excel(p, sheet_name="試合基本情報")
    except Exception:  # noqa: BLE001
        return set()
    return {str(int(g)) for g in info["試合ID"].dropna()}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", default="2026")
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    stale = []   # (league, game_type, date, 重複元の日付)
    partial = []
    for lg, lv in LEVELS:
        lv_dir = Path(npb.BASE_DATA_DIR) / f"{args.year}年" / lv
        if not lv_dir.is_dir():
            continue
        for type_dir in sorted(p for p in lv_dir.iterdir() if p.is_dir()):
            raw_root = type_dir / "raw"
            if not raw_root.is_dir():
                continue
            seen: dict = {}   # 試合ID -> 最初に保存された日付
            for d in sorted(p.name for p in raw_root.iterdir() if p.is_dir()):
                ids = game_ids_of(raw_root / d, d)
                if not ids:
                    continue
                dup = {g for g in ids if g in seen}
                if dup == ids:
                    stale.append((lg, type_dir.name, d, sorted({seen[g] for g in ids})))
                    continue            # 重複日付は「保存された日」として数えない
                if dup:
                    partial.append((lv, type_dir.name, d, len(dup), len(ids)))
                for g in ids:
                    seen.setdefault(g, d)

    print(f"重複日付（全試合が前の日付と同じ）: {len(stale)}件")
    for lg, gt, d, src in stale:
        print(f"  {'1軍' if lg == 'ichi' else '2軍'}/{gt} {d}  ← {', '.join(src)} と同じ試合")
    for lv, gt, d, n, total in partial:
        print(f"[要確認] {lv}/{gt} {d}: {total}試合中{n}試合が前の日付と重複（削除しない）")

    if not args.apply:
        print("\n（一覧のみ。--apply で削除）")
        return
    for lg, gt, d, _ in stale:
        npb.set_league_dirs(lg, d, gt)
        raw = Path(npb.RAW_DIR)
        if raw.is_dir():
            shutil.rmtree(raw)
            print(f"削除: {raw}")
        npb.remove_stale_date_outputs(d)


if __name__ == "__main__":
    main()
