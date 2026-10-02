"""保存済みのNPB投球データ（daily_pitch_data_*.xlsx）から、同じ投球の重複行を取り除く（通信なし）。

重複の原因：Yahoo!の打席ページは、走者の動き・牽制などでページが増えるたびに「その打席のここまでの全投球」を
表に並べ直すため、取得処理がそれを毎回取り込んでいた（取得処理は修正済み）。
キーはデータマート（run.py の preprocess_pitch）と同じ「試合ID＋投手名＋通算投球数」、最初の1件を残す。
→ データマート・数値JSONは元々この除去を通っているので、除去後に再計算しても結果は変わらない（RAWを正すためのもの）。

  python baseball/scripts/dedupe_pitch_data.py --year 2026            # 書き換える
  python baseball/scripts/dedupe_pitch_data.py --year 2026 --dry-run  # 件数だけ表示
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

_SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DATA_DIR = _SCRIPT_DIR.parent.parent / "data" / "baseball" / "プロ野球"
KEYS = ["試合ID", "投手名", "通算投球数"]   # run.py の PITCH_DEDUP_KEYS と同じ


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", default="2026")
    ap.add_argument("--base", default=str(BASE_DATA_DIR))
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    files = sorted(Path(args.base).glob(f"{args.year}年/*/*/raw/*/daily_pitch_data_*.xlsx"))
    print(f"対象ファイル: {len(files)}件")
    changed = removed_total = 0
    for f in files:
        try:
            df = pd.read_excel(f, engine="openpyxl")
        except Exception as e:  # noqa: BLE001
            print(f"[WARN] 読めない: {f} ({e})")
            continue
        if not all(c in df.columns for c in KEYS):
            continue
        new = df.drop_duplicates(subset=KEYS, keep="first").reset_index(drop=True)
        n = len(df) - len(new)
        if n <= 0:
            continue
        changed += 1
        removed_total += n
        print(f"{f.relative_to(args.base)}: {len(df)} → {len(new)}球（重複{n}）")
        if not args.dry_run:
            new.to_excel(f, index=False, engine="openpyxl")
    print(f"\n重複を除いたファイル: {changed}件 / 除いた行: {removed_total}行" + ("（dry-run: 書き換えなし）" if args.dry_run else ""))


if __name__ == "__main__":
    sys.exit(main())
