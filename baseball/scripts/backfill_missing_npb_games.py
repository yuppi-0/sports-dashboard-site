"""日程ページにあるのにRAWに無い試合（取得が丸ごと抜けた日・中止後の振替試合）を自動で取り直す。

  python baseball/scripts/backfill_missing_npb_games.py 2026-10-03 [--days 45] [--max-dates 10] [--dry-run]

対象日の直近 --correction-days 日の保存済み試合は、現在のYahoo!のページと比べ、後日の公式記録の訂正（安打↔失策の判定変更など）が
未反映なら、その日を取り直す。

対象日の直近 --days 日について find_incomplete_npb_games.py --schedule で「日程にあるのにRAWに無い」試合を探し、
該当する 日付×リーグ ごとに run.py（games pitch datamart）を実行する。取りこぼしは翌日以降の自動更新で自己修復される。
未完了試合の再取得までは行わない（コールドゲーム等の誤検知で毎日取り直しになるのを避けるため）。

実行時間を抑えるための仕組み（以前は取り直しを毎回10件×2軍まで走らせ、数十分〜数時間かかることがあった）:
  ・時間予算（--budget-sec、既定600秒）を超えたら新しい取り直しを始めない（残りは次回以降）
  ・取り直しても生成されない日（中止などで試合が無い）を記録（_backfill_state.json）し、2回試して出来なければ諦める＝毎日取り直し続けない
  ・1回の取り直しにも時間制限（--run-timeout、既定300秒）
  ・1軍を優先し、新しい日付から取り直す
"""
from __future__ import annotations
import argparse
import datetime
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
MARK = "日程にあるのにRAWに無い"
DEFAULT_BASE = Path("data/baseball/プロ野球")
MAX_TRIES = 2          # 取り直しても生成されなかった日は、これだけ試したら諦める
RETRY_AFTER_H = 20     # 同じ日を取り直す最短間隔（時間）


def load_state(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")


def raw_exists(base: Path, year: int, level: str, date: str) -> bool:
    """その日のRAW（all_gamesのxlsx）が、そのリーグのどこかの大会フォルダにあるか。"""
    return any((base / f"{year}年" / level).glob(f"*/raw/{date}/all_games_{date}.xlsx"))


def eligible(state: dict, level: str, date: str, now: float) -> bool:
    st = state.get(f"{level}|{date}")
    if not st:
        return True
    if st.get("tries", 0) >= MAX_TRIES:
        return False
    return (now - st.get("last", 0)) >= RETRY_AFTER_H * 3600


def find_missing(target: str, days: int, timeout: int = 300) -> list:
    """[(リーグ, 日付)] を返す。"""
    end = datetime.date.fromisoformat(target.split(":")[-1])
    start = end - datetime.timedelta(days=days)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "r.json"
        try:
            subprocess.run([sys.executable, str(HERE / "find_incomplete_npb_games.py"), "--year", str(end.year),
                            "--schedule", "--date", f"{start}:{end}", "--out", str(out)], check=False, timeout=timeout)
        except subprocess.TimeoutExpired:
            print("::warning::日程との突き合わせが時間切れになりました（次回以降に持ち越し）")
        if not out.exists():
            return []
        data = json.loads(out.read_text(encoding="utf-8"))
    need = {(p["level"], p["date"]) for p in data.get("problems", []) if any(MARK in x for x in p.get("problems", []))}
    return sorted(need, key=lambda x: (x[1], x[0]))


def find_corrected(target: str, days: int) -> list:
    """直近 days 日の保存済み1軍試合のうち、現在のYahoo!のページと違う（後日の公式記録の訂正が未反映の）日付を返す。"""
    sys.path.insert(0, str(HERE))
    import check_npb_stale_games as stale  # noqa: WPS433
    end = datetime.date.fromisoformat(target.split(":")[-1])
    start = end - datetime.timedelta(days=days)
    root = Path(stale.BASE) / f"{end.year}年" / "1軍" / "レギュラーシーズン" / "raw"
    return stale.changed_dates(root, start.isoformat(), end.isoformat())


def run_one(level: str, date: str, timeout: int) -> bool:
    flag = "--1軍" if level == "1軍" else "--2軍"
    cmd = [sys.executable, str(HERE / "run.py"), "--date", date, "--steps", "games", "pitch", "datamart", flag]
    print("実行:", " ".join(cmd), flush=True)
    try:
        subprocess.run(cmd, check=False, timeout=timeout)
        return True
    except subprocess.TimeoutExpired:
        print(f"::warning::{level} {date} の取り直しが時間切れになりました")
        return False


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target")
    ap.add_argument("--days", type=int, default=21)
    ap.add_argument("--max-dates", type=int, default=4, help="1回の実行で取り直す 日付×リーグ の上限")
    ap.add_argument("--correction-days", type=int, default=7, help="この日数以内の保存済み試合を現在のページと比べ、公式記録の訂正があれば取り直す（0で無効）")
    ap.add_argument("--budget-sec", type=int, default=600, help="この秒数を超えたら新しい取り直しを始めない")
    ap.add_argument("--run-timeout", type=int, default=300, help="1回の取り直し（run.py）の制限時間（秒）")
    ap.add_argument("--base", default=str(DEFAULT_BASE))
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    t0 = time.monotonic()
    base = Path(args.base)
    state_path = base / "_backfill_state.json"
    state = load_state(state_path)
    year = datetime.date.fromisoformat(args.target.split(":")[-1]).year
    now = time.time()

    def over_budget() -> bool:
        return time.monotonic() - t0 > args.budget_sec

    found = find_missing(args.target, args.days, timeout=min(300, args.budget_sec))
    print(f"取りこぼし（日程にあるのにRAWに無い）: {len(found)}件 {found}")
    # 出来なかった日を諦める（取り直しても生成されない日＝中止などは、毎日取り直し続けない）
    gave_up = [(lv, d) for lv, d in found if not eligible(state, lv, d, now) and state.get(f"{lv}|{d}", {}).get("tries", 0) >= MAX_TRIES]
    if gave_up:
        print(f"取り直しても生成されなかったため諦め済み: {len(gave_up)}件 {gave_up}")
    need = [(lv, d) for lv, d in found if eligible(state, lv, d, now)]
    need.sort(key=lambda x: (0 if x[0] == "1軍" else 1, [-ord(c) for c in x[1]]))   # 1軍優先・新しい日付から
    if len(need) > args.max_dates:
        print(f"::warning::取りこぼしが多いため先頭{args.max_dates}件だけ取り直します（残りは次回以降）")
        need = need[:args.max_dates]
    done = []
    for lv, d in need:
        if over_budget():
            print(f"::warning::時間予算（{args.budget_sec}秒）を超えたため、残りの取り直しは次回以降に回します")
            break
        if not args.dry_run:
            run_one(lv, d, args.run_timeout)
            key = f"{lv}|{d}"
            if raw_exists(base, year, lv, d):
                state.pop(key, None)                     # 取れた
            else:
                st = state.setdefault(key, {"tries": 0})
                st["tries"] += 1
                st["last"] = now
            save_state(state_path, state)
        done.append((lv, d))
    if args.correction_days > 0 and not over_budget():
        corrected = find_corrected(args.target, args.correction_days)
        print(f"公式記録の訂正が未反映の日（1軍）: {len(corrected)}件 {corrected}")
        for d in corrected[:args.max_dates]:
            if over_budget():
                print("::warning::時間予算を超えたため、訂正の取り直しは次回以降に回します")
                break
            if not args.dry_run:
                run_one("1軍", d, args.run_timeout)
        if corrected:
            print(f"::warning::公式記録の訂正が未反映だった日を取り直しました: {', '.join(corrected)}")
    if done:
        print(f"::warning::取りこぼした試合を自動で取り直しました: {', '.join(f'{lv} {d}' for lv, d in done)}")
    print(f"所要 {int(time.monotonic() - t0)}秒")


if __name__ == "__main__":
    main()
