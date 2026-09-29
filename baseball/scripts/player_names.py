"""MLBAM ID → 選手名（"last, first" 小文字）の解決ヘルパー（並列実行対応）。

[背景] 以前は pybaseball.playerid_reverse_lookup を各プロセスが直接呼んでいた。
この関数は初回に Chadwick Register の zip をダウンロードして pybaseball 側のキャッシュに
書き込むが、複数プロセスを同時に走らせると
  - 全プロセスが同時にダウンロードして
  - 書きかけのキャッシュファイルを別プロセスが読んで "File is not a zip file" で落ちる
という競合が起きる。さらに解決に失敗すると打者名が選手ID（"657675"など）になり、
JSON/xlsx が壊れた名前で出力されてしまう。

[方針]
  1. ローカルのキャッシュ（JSON）を最優先で読む（読むだけならロック不要・追記は原子的に置換）
  2. 未解決IDがあるときだけ、ファイルロックを取って Chadwick Register を1回だけ取り込み直す
     （ロック取得後にもう一度キャッシュを確認するので、並列ジョブでもダウンロードは1回）
       a. raw.githubusercontent.com の people-*.csv（zipではなくCSV。zip系のプロキシ制限を避けられる）
       b. pybaseball.playerid_reverse_lookup（従来経路）
  3. それでも無いIDは MLB StatsAPI の people エンドポイントで補完（デビュー直後の新人など）
  4. 全部失敗したら呼び出し側が従来通りIDをフォールバック表示にする
"""
from __future__ import annotations
import contextlib
import io
import json
import os
import time

import pandas as pd
import requests

try:
    import fcntl  # POSIX のみ。Windowsでは排他ロック無しで動く（単独実行前提）
except ImportError:  # pragma: no cover
    fcntl = None

_REGISTER_URL = "https://raw.githubusercontent.com/chadwickbureau/register/master/data/people-{}.csv"
_REGISTER_SUFFIXES = "0123456789abcdef"
_REFRESH_INTERVAL_SEC = 12 * 3600   # 未解決IDがあっても、この間隔より短い間隔ではRegisterを取り直さない
_REFRESHED_KEY = "_refreshed_at"

_mem_cache: dict[str, dict] = {}


@contextlib.contextmanager
def _file_lock(path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    f = open(path + ".lock", "w")
    try:
        if fcntl is not None:
            fcntl.flock(f, fcntl.LOCK_EX)
        yield
    finally:
        try:
            if fcntl is not None:
                fcntl.flock(f, fcntl.LOCK_UN)
        finally:
            f.close()


def _load(cache_path: str) -> dict:
    try:
        with open(cache_path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save(cache_path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    tmp = f"{cache_path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    os.replace(tmp, cache_path)   # 原子的に差し替え（読み手が書きかけを見ない）


def _fmt(last, first) -> str:
    last = "" if pd.isna(last) else str(last).strip().lower()
    first = "" if pd.isna(first) else str(first).strip().lower()
    return f"{last}, {first}" if first else last


def _fetch_register_csv() -> dict:
    """Chadwick Register の people-*.csv を取り込み {mlbam_id: "last, first"} を返す。"""
    out: dict[str, str] = {}
    for s in _REGISTER_SUFFIXES:
        last_err = None
        for attempt in range(3):
            try:
                res = requests.get(_REGISTER_URL.format(s), timeout=60)
                res.raise_for_status()
                df = pd.read_csv(io.BytesIO(res.content), low_memory=False,
                                 usecols=["key_mlbam", "name_last", "name_first"])
                break
            except Exception as e:  # noqa: BLE001
                last_err = e
                time.sleep(2 * (attempt + 1))
        else:
            raise RuntimeError(f"Chadwick people-{s}.csv 取得失敗: {last_err}")
        df = df.dropna(subset=["key_mlbam"])
        for mid, last, first in zip(df["key_mlbam"], df["name_last"], df["name_first"]):
            name = _fmt(last, first)
            if name:
                out[str(int(mid))] = name
    return out


def _fetch_pybaseball(ids: list[int]) -> dict:
    from pybaseball import playerid_reverse_lookup
    out: dict[str, str] = {}
    lookup = playerid_reverse_lookup(ids, key_type="mlbam")
    for _, row in lookup.iterrows():
        mid = row.get("key_mlbam")
        name = _fmt(row.get("name_last"), row.get("name_first"))
        if pd.notna(mid) and name:
            out[str(int(mid))] = name
    return out


def _fetch_statsapi(ids: list[int]) -> dict:
    out: dict[str, str] = {}
    for i in range(0, len(ids), 100):
        chunk = ids[i:i + 100]
        try:
            res = requests.get("https://statsapi.mlb.com/api/v1/people",
                               params={"personIds": ",".join(map(str, chunk))}, timeout=30)
            res.raise_for_status()
            for p in res.json().get("people", []):
                name = _fmt(p.get("lastName"), p.get("firstName") or p.get("useName"))
                if name:
                    out[str(p["id"])] = name
        except Exception as e:  # noqa: BLE001
            print(f"  [WARN] StatsAPI選手名取得失敗: {e}")
    return out


def resolve_names(ids, cache_path: str) -> dict[int, str]:
    """MLBAM ID のリストを {id: "last, first"} に解決する（解決できなかったIDは含まれない）。"""
    ids = sorted({int(i) for i in ids})
    if not ids:
        return {}

    cache = _mem_cache.get(cache_path)
    if cache is None:
        cache = _mem_cache[cache_path] = _load(cache_path)
    missing = [i for i in ids if str(i) not in cache]

    if missing:
        with _file_lock(cache_path):
            cache = _load(cache_path)               # 他プロセスが更新済みかもしれないので読み直す
            missing = [i for i in ids if str(i) not in cache]
            stale = time.time() - float(cache.get(_REFRESHED_KEY, 0)) > _REFRESH_INTERVAL_SEC
            if missing and stale:
                print(f"  選手名キャッシュに無い {len(missing)}名 → Chadwick Registerを取り込みます")
                fetched: dict = {}
                try:
                    fetched = _fetch_register_csv()
                except Exception as e:  # noqa: BLE001
                    print(f"  [WARN] Register(CSV)取得失敗: {e} → pybaseballで再試行")
                    try:
                        fetched = _fetch_pybaseball(missing)
                    except Exception as e2:  # noqa: BLE001
                        print(f"  [WARN] pybaseball lookup失敗: {e2}")
                cache.update(fetched)
                if fetched:
                    cache[_REFRESHED_KEY] = time.time()
                missing = [i for i in ids if str(i) not in cache]
            if missing:
                got = _fetch_statsapi(missing)
                cache.update(got)
                missing = [i for i in ids if str(i) not in cache]
            _save(cache_path, cache)
            if missing:
                print(f"  [WARN] 選手名を解決できないID: {len(missing)}件（IDのまま表示します）")
        _mem_cache[cache_path] = cache

    return {i: cache[str(i)] for i in ids if str(i) in cache}
