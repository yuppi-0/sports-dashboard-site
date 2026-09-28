"""既存の JSON を一括で .json.gz にする（docs/ 容量対策）。

対象（これ以外は触らない）:
  - games/json/YYYY-MM-DD.json            （日別）
  - pitcher_cards_numeric/*.json / batter_cards_numeric/*.json（index.json は除く）
  - batter_cards_numeric/pivot_population_*.json
安全策: 圧縮 → 展開して元と完全一致を検証 → 一致した時だけ元ファイルを削除。--apply を付けない限りドライラン。

使い方:
  python compress_existing.py docs/data --dry-run        # 件数とサイズ見積りだけ
  python compress_existing.py docs/data --apply          # 実行
"""
import argparse, gzip, io, os, re, sys

DAILY = re.compile(r"^\d{4}-\d{2}-\d{2}\.json$")
PIVOT = re.compile(r"^pivot_population_\d+\.json$")


def is_target(root: str, name: str) -> bool:
    parent = os.path.basename(root)
    if name == "index.json":
        return False
    if parent == "json" and os.path.basename(os.path.dirname(root)) == "games":
        return bool(DAILY.match(name))
    if parent in ("pitcher_cards_numeric", "batter_cards_numeric"):
        return name.endswith(".json") and (not name.startswith("pivot_") or bool(PIVOT.match(name)))
    return False


def gz_bytes(raw: bytes) -> bytes:
    buf = io.BytesIO()
    with gzip.GzipFile(filename="", mode="wb", fileobj=buf, compresslevel=9, mtime=0) as g:
        g.write(raw)
    return buf.getvalue()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    n = before = after = 0
    for root, _, files in os.walk(a.root):
        for name in files:
            if not is_target(root, name):
                continue
            src = os.path.join(root, name)
            raw = open(src, "rb").read()
            comp = gz_bytes(raw)
            n += 1; before += len(raw); after += len(comp)
            if not a.apply:
                continue
            dst = src + ".gz"
            with open(dst + ".tmp", "wb") as f:
                f.write(comp)
            if gzip.decompress(open(dst + ".tmp", "rb").read()) != raw:
                os.remove(dst + ".tmp"); print(f"[NG] 検証不一致のためスキップ: {src}", file=sys.stderr); continue
            os.replace(dst + ".tmp", dst)
            os.remove(src)
    print(f"{'実行' if a.apply else 'ドライラン'}: {n}ファイル  {before/1e6:.0f}MB → {after/1e6:.0f}MB")


if __name__ == "__main__":
    main()
