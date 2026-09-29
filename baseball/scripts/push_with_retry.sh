#!/usr/bin/env bash
# 並列ジョブ（日付範囲を分けた複数のworkflow_dispatch）が同じブランチへ同時にpushしても
# 取りこぼさないためのpushヘルパー。カレントディレクトリのリポジトリを対象にする。
#
#   push_with_retry.sh [REBUILD_CMD]
#
# 従来の「git push || git pull --rebase」のループには、次の問題があった:
#   - 並列ジョブ同士がindex.jsonなど共有テキストファイルを書き換えているとrebaseが競合で止まり、
#     以降のpushが全部失敗する（ループ終了後もステップが成功扱いになり、データが黙って消える）
# このスクリプトは
#   1. rebase中断状態を必ず解除してから取り込み直す
#   2. 競合は「自分のコミット側」を優先して自動解決（日別ファイルはジョブ間で被らない）
#   3. 共有のindex.jsonはREBUILD_CMD（rebuild_index.py）で実在ファイルから再生成してコミット
#   4. 最終的にpushできなければ非0で終了（ステップを失敗にする）
# 現在のブランチ（detachedならGITHUB_REF_NAME）へpushする。
set -u
REBUILD_CMD="${1:-}"
# git add の対象（サイト側リポジトリでは、ネスト展開したdata/を巻き込まないよう ADD_PATHS=docs を指定する）
ADD_PATHS="${ADD_PATHS:-.}"
BRANCH="$(git rev-parse --abbrev-ref HEAD)"
[ "$BRANCH" = "HEAD" ] && BRANCH="${GITHUB_REF_NAME:-main}"

for i in $(seq 1 10); do
  if git push origin "HEAD:refs/heads/$BRANCH"; then
    exit 0
  fi
  echo "push失敗。リモートの変更を取り込み直します (試行 $i/10)"
  git rebase --abort 2>/dev/null || true
  git fetch origin "$BRANCH" || { sleep 5; continue; }
  if ! git rebase -X theirs "origin/$BRANCH"; then
    # -X theirs（rebase中は「再適用する自分のコミット」側）で解決できなかった競合
    # （バイナリ・削除vs変更など）は、自分側を採用して続行する
    git checkout --theirs -- $ADD_PATHS 2>/dev/null || true
    git add -A -- $ADD_PATHS
    if ! GIT_EDITOR=true git rebase --continue; then
      git rebase --abort 2>/dev/null || true
      sleep $((RANDOM % 10 + 3))
      continue
    fi
  fi
  if [ -n "$REBUILD_CMD" ]; then
    eval "$REBUILD_CMD"
    git add -A -- $ADD_PATHS
    if ! git diff --cached --quiet; then
      git commit -q -m "index.json を再生成（並列実行の競合解消）"
    fi
  fi
  # 並列ジョブ同士のpushが同時に再試行しないよう、ランダムに待つ
  sleep $((RANDOM % 10 + 3))
done
echo "::error::push を10回試行しても成功しませんでした"
exit 1
