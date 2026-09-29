"""run_mlb.py を日付範囲で分割して並列実行するランチャー。

  python run_mlb_parallel.py --date 2024-03-20:2024-11-05 --steps datamart -j 4
  python run_mlb_parallel.py --date 2023-03-30:2023-11-05 --steps games -j 4
  python run_mlb_parallel.py --steps defense --year 2024 2025 2026

動作:
  1. 日付範囲を -j 個のチャンクに均等分割し、run_mlb.py を別プロセスで同時実行する。
     各プロセスは --skip-llm-input --skip-batter-llm-input 付き（シーズン共有ファイルを触らない）。
     日別ファイル（raw/datamart/json.gz）はプロセス間で衝突せず、index.json はロック＋原子置換、
     選手名キャッシュもロック付きなので並列でも壊れない。
  2. 全チャンク完了後、シーズン共有の llm_input / batter_llm_input を年ごとに1回だけ実行する
     （--no-finalize で省略可）。
  3. 1つでも失敗したチャンクがあれば非0で終了（失敗チャンクの日付範囲を表示）。
"""
from __future__ import annotations
import argparse
import concurrent.futures
import datetime
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
RUN = str(HERE / "run_mlb.py")


def split_dates(date_arg: str, n: int) -> list[str]:
    if ":" not in date_arg:
        return [date_arg]
    a, b = (datetime.date.fromisoformat(x.strip()) for x in date_arg.split(":", 1))
    days = [(a + datetime.timedelta(days=i)) for i in range((b - a).days + 1)]
    n = max(1, min(n, len(days)))
    size = -(-len(days) // n)   # 切り上げ
    return [f"{days[i].isoformat()}:{days[min(i + size, len(days)) - 1].isoformat()}"
            for i in range(0, len(days), size)]


def run_chunk(idx: int, chunk: str, steps: list[str], extra: list[str]) -> tuple[str, int]:
    log = HERE / f".parallel_{idx}.log"
    cmd = [sys.executable, RUN, "--date", chunk, "--steps", *steps,
           "--skip-llm-input", "--skip-batter-llm-input", *extra]
    with open(log, "w", encoding="utf-8") as f:
        rc = subprocess.call(cmd, stdout=f, stderr=subprocess.STDOUT)
    print(f"[chunk {idx}] {chunk} → {'OK' if rc == 0 else f'FAILED(rc={rc})'}  log: {log.name}", flush=True)
    return chunk, rc


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--date", help="単日 or 開始日:終了日")
    ap.add_argument("--steps", nargs="+", default=["datamart"])
    ap.add_argument("--year", nargs="+", type=int, help="--steps defense の対象年（複数可）")
    ap.add_argument("-j", "--jobs", type=int, default=4)
    ap.add_argument("--no-finalize", action="store_true", help="llm_input/batter_llm_input の最終再生成をしない")
    ap.add_argument("--refetch", action="store_true")
    args, _ = ap.parse_known_args()

    if "defense" in args.steps:
        cmd = [sys.executable, RUN, "--steps", "defense", "--date", args.date or f"{args.year[0]}-01-01",
               "--year", *map(str, args.year or [])]
        return subprocess.call(cmd)

    chunks = split_dates(args.date, args.jobs)
    extra = ["--refetch"] if args.refetch else []
    print(f"{len(chunks)}チャンクを並列実行: {chunks}", flush=True)
    failed = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(chunks)) as ex:
        for chunk, rc in ex.map(lambda t: run_chunk(t[0], t[1], args.steps, extra), enumerate(chunks)):
            if rc != 0:
                failed.append(chunk)

    if not args.no_finalize:
        end = chunks[-1].split(":")[-1]
        for step in ("llm_input", "batter_llm_input"):
            print(f"[finalize] --steps {step} --date {end}", flush=True)
            rc = subprocess.call([sys.executable, RUN, "--steps", step, "--date", end])
            if rc != 0:
                failed.append(step)

    if failed:
        print(f"失敗: {failed}", file=sys.stderr)
        return 1
    print("並列実行 完了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
