"""MLB公式日程（statsapi）で終了しているのに当サイトのデータに無い試合を、自動で取り直す。

  python baseball/scripts/backfill_missing_mlb_games.py 2026-10-03 [--days 20] [--max-dates 8] [--dry-run]

Statcastの反映が遅れた日にRAWキャッシュが作られると、その日の一部の試合が欠けたまま固定される
（キャッシュがあると再取得しないため）。対象日の直近 --days 日について公式日程と突き合わせ、欠けた試合のある日を
run_mlb.py --steps games --refetch で取り直す。日付ラベルは日本時間基準（米国日付 + 1日）。
"""
from __future__ import annotations

import argparse
import datetime
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_mlb_schedule_gaps as gap  # noqa: E402

FINAL = ("Final", "Game Over", "Completed Early")


def missing_labels(target: str, days: int) -> list:
    end = datetime.date.fromisoformat(target.split(":")[-1])
    start = end - datetime.timedelta(days=days)
    ours = gap.our_games(str(end.year))
    labels = set()
    for g in gap.official(str(end.year)):
        if g["state"] not in FINAL or g["pk"] in ours:
            continue
        d = datetime.date.fromisoformat(g["date"])
        if start <= d <= end:
            labels.add((d + datetime.timedelta(days=1)).isoformat())
    return sorted(labels)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target")
    ap.add_argument("--days", type=int, default=1, help="公式日程と突き合わせる日数（毎日実行しているので、前日で足りる。広く調べたいときは手動実行で増やす）")
    ap.add_argument("--max-dates", type=int, default=4)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    labels = missing_labels(args.target, args.days)
    print(f"MLB取りこぼし（公式で終了済みなのに無い試合のある日）: {len(labels)}件 {labels}")
    if len(labels) > args.max_dates:
        print(f"::warning::取りこぼしが多いため先頭{args.max_dates}件だけ取り直します（残りは次回以降）")
        labels = labels[:args.max_dates]
    for d in labels:
        cmd = [sys.executable, str(HERE / "run_mlb.py"), "--date", d, "--steps", "games", "--refetch"]
        print("実行:", " ".join(cmd))
        if not args.dry_run:
            subprocess.run(cmd, check=False)
    if labels:
        print(f"::warning::MLBの取りこぼした試合を自動で取り直しました: {', '.join(labels)}")


if __name__ == "__main__":
    main()
