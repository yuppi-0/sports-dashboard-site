"""点検（find_incomplete_npb_games.py）で問題が見つかった試合を、その試合だけ自動で取り直す。
日程ページにあるのにRAWに無い試合（取得が丸ごと抜けた日・中止後の振替試合）も含む。

  python baseball/scripts/backfill_missing_npb_games.py 2026-10-03 [--days 45] [--max-dates 10] [--dry-run]

対象日の直近 --correction-days 日の保存済み試合は、現在のYahoo!のページと比べ、後日の公式記録の訂正（安打↔失策の判定変更など）が
未反映なら、その日を取り直す。

対象日の直近 --days 日について find_incomplete_npb_games.py --schedule で「日程にあるのにRAWに無い」試合を探し、
該当する 日付×リーグ ごとに run.py（games pitch datamart）を実行する。取りこぼしは翌日以降の自動更新で自己修復される。
点検で問題が出た試合は、取り直す（試合状態が終了でない・成績やスコアボードが無い・投球データの有無／最終回／過不足／重複・スコア不一致など）。
取り直しは前日分だけを1回なので、コールドゲーム等で同じ問題が出続ける試合も、取り直しは1日1回で済む。

実行時間を抑えるための仕組み（以前は取り直しを毎回10件×2軍まで走らせ、数十分〜数時間かかることがあった）:
  ・時間予算（--budget-sec、既定は制限なし）。毎日の対象は前日分の問題のある試合だけで少ない。手動で広く調べるときに指定できる
  ・取り直しても生成されない日（中止などで試合が無い）を記録（_backfill_state.json）し、1回試して直らなければ諦める＝毎日取り直し続けない
  ・1回の取り直しにも時間制限（--run-timeout、既定1800秒。その日を丸ごと取り直すと10分以上かかることがある）
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
MAX_TRIES = 1          # 取り直しは1回だけ（もう一度取り直しても変わらないので、再試行しない）
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


def check_range(target: str, days: int) -> tuple:
    """確認する日付の範囲 (開始, 終了)。対象日の「前日」から days 日ぶんさかのぼる。
    対象日（当日）は、この実行の主な取得手順がちょうど取っている日なので含めない（翌日の実行が前日として確認する）。"""
    last = datetime.date.fromisoformat(target.split(":")[-1]) - datetime.timedelta(days=1)
    return last - datetime.timedelta(days=max(days, 1) - 1), last


def find_missing(target: str, days: int, timeout: int = 300) -> list:
    """[(リーグ, 日付)] を返す。"""
    start, end = check_range(target, days)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "r.json"
        try:
            subprocess.run([sys.executable, str(HERE / "find_incomplete_npb_games.py"), "--year", str(end.year),
                            "--schedule", "--date", f"{start}:{end}", "--gh-warning", "--out", str(out)], check=False, timeout=timeout)
        except subprocess.TimeoutExpired:
            print("::warning::日程との突き合わせが時間切れになりました（次回以降に持ち越し）")
        if not out.exists():
            return []
        data = json.loads(out.read_text(encoding="utf-8"))
    return group_problems(data.get("problems", []))


def group_problems(problems: list) -> list:
    """点検で見つかった問題を、(リーグ, 日付) ごとに [(リーグ, 日付, [問題のある試合ID], [理由])] にまとめる。
    試合IDがある問題は、その試合だけを取り直す対象にする（日程にあるのにRAWが無い・最後まで取れていない・投球データ不足や重複・
    スコア不一致・成績が無い など、点検が出す問題すべて）。
    試合IDが無い問題（その日の all_games が無い・別の日付と同じ試合を保存 など）は、試合を特定できないので取り直さず、警告に出すだけ。
    とくに「all_games無し」は、2軍のシーズン終了後や試合の無い日のフォルダにも出る（2026年は10月2〜4日の2軍など）ので、取り直しても意味がない。
    本当に取りこぼした試合は、日程との突き合わせ（日程にあるのにRAWに無い）が試合IDつきで拾う。"""
    groups: dict = {}
    for p in problems:
        if p.get("cancelled"):
            continue
        key = (p["level"], p["date"])
        g = groups.setdefault(key, {"gids": [], "reasons": []})
        reasons = p.get("problems", [])
        if p.get("gid"):
            gid = str(p["gid"])
            if gid not in g["gids"]:
                g["gids"].append(gid)
            g["reasons"] += [f"{gid}: {x}" for x in reasons]
        else:
            print(f"::warning::取り直せない問題（試合を特定できない）: {key[0]} {key[1]} {' / '.join(reasons)}")
    out = []
    for (lv, d), g in sorted(groups.items(), key=lambda x: (x[0][1], x[0][0])):
        if g["gids"]:
            out.append((lv, d, g["gids"], g["reasons"]))
    return out


def find_corrected(target: str, days: int) -> list:
    """直近 days 日の保存済み1軍試合のうち、現在のYahoo!のページと違う（後日の公式記録の訂正が未反映の）日付を返す。"""
    sys.path.insert(0, str(HERE))
    import check_npb_stale_games as stale  # noqa: WPS433
    start, end = check_range(target, days)
    root = Path(stale.BASE) / f"{end.year}年" / "1軍" / "レギュラーシーズン" / "raw"
    return stale.changed_dates(root, start.isoformat(), end.isoformat())


def run_one(level: str, date: str, timeout: int, game_ids: list | None = None) -> bool:
    """その日の games・pitch・datamart を取り直す。game_ids があれば、その試合だけ取得する（日全体は取り直さない）。"""
    flag = "--1軍" if level == "1軍" else "--2軍"
    cmd = [sys.executable, str(HERE / "run.py"), "--date", date, "--steps", "games", "pitch", "datamart", *(game_ids or []), flag]
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
    ap.add_argument("--days", type=int, default=1, help="確認する日数（対象日の前日からさかのぼる。1＝前日だけ。毎日実行しているので1で足りる。広く調べたいときは手動実行で増やす）")
    ap.add_argument("--max-dates", type=int, default=4, help="1回の実行で取り直す 日付×リーグ の上限")
    ap.add_argument("--correction-days", type=int, default=0, help="対象日の前日からこの日数ぶんの保存済み試合を現在のページと比べ、公式記録の訂正があれば取り直す（0で無効＝既定）。公式記録の訂正は前日のうちには入らないことが多いので、毎日の実行では見ない。調べたいときは手動実行で日数を指定する")
    ap.add_argument("--budget-sec", type=int, default=0, help="この秒数を超えたら新しい取り直しを始めない（0＝制限なし。毎日の件数は前日分の問題のある試合だけで少ないので既定は制限しない）")
    ap.add_argument("--run-timeout", type=int, default=1800, help="1回の取り直し（run.py）の制限時間（秒）。その日を丸ごと取り直すと10分以上かかることがあるので長めにしてある")
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
        return args.budget_sec > 0 and time.monotonic() - t0 > args.budget_sec

    t_phase = time.monotonic()
    found4 = find_missing(args.target, args.days, timeout=1800)
    print(f"[所要] 点検（日程との突き合わせ含む） {int(time.monotonic() - t_phase)}秒")
    gids_of = {(lv, d): g for lv, d, g, _r in found4}
    found = [(lv, d) for lv, d, _g, _r in found4]
    print(f"取り直す対象（点検で問題があった試合）: {len(found)}件")
    for lv, d, g, reasons in found4:
        print(f"  {lv} {d} 試合ID={g or '（その日を丸ごと）'} :: {' / '.join(reasons)[:300]}")
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
            t_run = time.monotonic()
            run_one(lv, d, args.run_timeout, gids_of.get((lv, d)))   # 問題があった試合だけ取得する
            print(f"[所要] 取り直し {lv} {d} {int(time.monotonic() - t_run)}秒")
            key = f"{lv}|{d}"
            st = state.setdefault(key, {"tries": 0})       # 試合単位の取り直しは直ったか分からないので、試した回数を数える（1回で打ち止め）
            st["tries"] += 1
            st["last"] = now
            save_state(state_path, state)
        done.append((lv, d))
    if args.correction_days > 0 and not over_budget():
        t_phase = time.monotonic()
        corrected = find_corrected(args.target, args.correction_days)
        print(f"[所要] 訂正の確認 {int(time.monotonic() - t_phase)}秒")
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
