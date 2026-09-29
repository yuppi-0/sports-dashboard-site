"""docs/ 配下の日別JSON一覧 index.json を、実在するファイルから再生成する。

並列実行した複数ジョブがそれぞれ index.json を書き換えて git 上で競合した後に、
「実在する YYYY-MM-DD.json(.gz) を全部列挙し直す」ことで正しい内容に戻すために使う
（push_with_retry.sh から自動で呼ばれる。手動実行も可）。

  python baseball/scripts/rebuild_index.py            # docs/baseball/data 配下すべて
  python baseball/scripts/rebuild_index.py <dir>...   # 指定した games/json ディレクトリだけ
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from jsonio import list_json_stems  # noqa: E402

ROOT = Path(__file__).resolve().parents[2] / "docs" / "baseball" / "data"


def rebuild(json_dir: str) -> bool:
    files = list_json_stems(json_dir)
    if not files:
        return False
    path = os.path.join(json_dir, "index.json")
    new = json.dumps({"files": files}, ensure_ascii=False, indent=2)
    try:
        with open(path, encoding="utf-8") as f:
            if f.read() == new:
                return False
    except FileNotFoundError:
        pass
    tmp = f"{path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(new)
    os.replace(tmp, path)
    print(f"再生成: {path} ({len(files)}件)")
    return True


def main() -> None:
    dirs = sys.argv[1:] or sorted(str(p) for p in ROOT.rglob("games/json") if p.is_dir())
    n = sum(rebuild(d) for d in dirs)
    print(f"index.json 再生成: {n}件")


if __name__ == "__main__":
    main()
