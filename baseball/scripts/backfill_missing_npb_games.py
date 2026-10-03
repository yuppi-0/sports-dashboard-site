"""日程ページにあるのにRAWに無い試合（取得が丸ごと抜けた日・中止後の振替試合）を自動で取り直す。

  python baseball/scripts/backfill_missing_npb_games.py 2026-10-03 [--days 45] [--max-dates 10] [--dry-run]

対象日の直近 --days 日について find_incomplete_npb_games.py --schedule で「日程にあるのにRAWに無い」試合を探し、
該当する 日付×リーグ ごとに run.py（games pitch datamart）を実行する。取りこぼしは翌日以降の自動更新で自己修復される。
未完了試合の再取得までは行わない（コールドゲーム等の誤検知で毎日取り直しになるのを避けるため）。
"""
from __future__ import annotations
import argparse
import datetime
import json
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
MARK = "日程にあるのにRAWに無い"


def find_missing(target: str, days: int) -> list:
    """[(リーグ, 日付)] を返す。"""
    end = datetime.date.fromisoformat(target.split(":")[-1])
    start = end - datetime.timedelta(days=days)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "r.json"
        subprocess.run([sys.executable, str(HERE / "find_incomplete_npb_games.py"), "--year", str(end.year),
                        "--schedule", "--date", f"{start}:{end}", "--out", str(out)], check=False)
        if not out.exists():
            return []
        data = json.loads(out.read_text(encoding="utf-8"))
    need = {(p["level"], p["date"]) for p in data.get("problems", []) if any(MARK in x for x in p.get("problems", []))}
    return sorted(need, key=lambda x: (x[1], x[0]))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target")
    ap.add_argument("--days", type=int, default=45)
    ap.add_argument("--max-dates", type=int, default=10)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    need = find_missing(args.target, args.days)
    print(f"取りこぼし（日程にあるのにRAWに無い）: {len(need)}件 {need}")
    if len(need) > args.max_dates:
        print(f"::warning::取りこぼしが多いため先頭{args.max_dates}件だけ取り直します（残りは次回以降）")
        need = need[:args.max_dates]
    for lv, d in need:
        flag = "--1軍" if lv == "1軍" else "--2軍"
        cmd = [sys.executable, str(HERE / "run.py"), "--date", d, "--steps", "games", "pitch", "datamart", flag]
        print("実行:", " ".join(cmd))
        if not args.dry_run:
            subprocess.run(cmd, check=False)
    if need:
        print(f"::warning::取りこぼした試合を自動で取り直しました: {', '.join(f'{lv} {d}' for lv, d in need)}")


if __name__ == "__main__":
    main()
