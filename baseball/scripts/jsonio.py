"""JSON の .json.gz 読み書き共通ヘルパー（docs/ 容量対策）。

- 書き込み: write_json(path, obj) → path が .json でも .json.gz として保存（mtime=0・level9 で決定的）
- 読み込み: read_json(path)       → .json.gz があればそれ、無ければ .json を読む
- 列挙:     list_json_stems(dir)  → 日別 JSON の「論理名」(YYYY-MM-DD.json) を .json/.json.gz 両対応で返す
index.json など小さいファイルは圧縮せず、従来どおり .json のまま。
"""
from __future__ import annotations
import gzip, io, json, os


def gz_path(path: str) -> str:
    return path if path.endswith(".gz") else path + ".gz"


def plain_path(path: str) -> str:
    return path[:-3] if path.endswith(".gz") else path


def write_json(path: str, obj, *, indent=None, separators=(",", ":"), ensure_ascii=False) -> str:
    """obj を gzip 圧縮して <path>.gz に保存し、同名の非圧縮 .json があれば削除する。"""
    out = gz_path(path)
    raw = json.dumps(obj, ensure_ascii=ensure_ascii, indent=indent,
                     separators=None if indent else separators).encode("utf-8")
    buf = io.BytesIO()
    # mtime=0・filename空 → 同じ内容なら同じバイト列（gitの差分が無駄に増えない）
    with gzip.GzipFile(filename="", mode="wb", fileobj=buf, compresslevel=9, mtime=0) as gz:
        gz.write(raw)
    tmp = out + ".tmp"
    with open(tmp, "wb") as f:
        f.write(buf.getvalue())
    os.replace(tmp, out)
    old = plain_path(path)
    if os.path.exists(old):
        os.remove(old)
    return out


def read_json(path: str):
    """<path>.gz があればそれを、無ければ <path>（.json）を読む。"""
    gz, pl = gz_path(path), plain_path(path)
    if os.path.exists(gz):
        with gzip.open(gz, "rt", encoding="utf-8") as f:
            return json.load(f)
    with open(pl, "r", encoding="utf-8") as f:
        return json.load(f)


def exists_json(path: str) -> bool:
    return os.path.exists(gz_path(path)) or os.path.exists(plain_path(path))


def remove_json(path: str) -> bool:
    removed = False
    for p in (gz_path(path), plain_path(path)):
        if os.path.exists(p):
            os.remove(p)
            removed = True
    return removed


def list_json_stems(json_dir: str) -> list[str]:
    """YYYY-MM-DD.json 形式の論理名を重複なしでソートして返す（index.json / season_* は除外）。"""
    names = set()
    for f in os.listdir(json_dir):
        n = plain_path(f)
        if n.endswith(".json") and n != "index.json" and not n.startswith("season_"):
            names.add(n)
    return sorted(names)


def prune_orphan_cards(json_dir: str, keep_ids, protect_prefixes=("pivot_population_",)) -> list[str]:
    """選手カードJSON置き場から、選手一覧(index.json)に載らない古いファイルを削除して、削除した論理名を返す。
    選手名の表記ゆれ・ID解決の失敗で出来た旧ファイル（例: josé_a_ferrer と jose_a_ferrer、数字だけの名前）が
    残り続けないようにする。index.json と protect_prefixes で始まる補助ファイルは消さない。"""
    keep = set(keep_ids)
    removed = []
    for f in os.listdir(json_dir):
        stem = plain_path(f)
        if not stem.endswith(".json") or stem == "index.json" or stem.startswith(tuple(protect_prefixes)):
            continue
        sid = stem[:-5]
        if sid not in keep:
            os.remove(os.path.join(json_dir, f))
            removed.append(sid)
    return sorted(set(removed))
