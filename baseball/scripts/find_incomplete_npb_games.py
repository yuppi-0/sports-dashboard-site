"""NPBの「最後まで取得できていない／取得に失敗している」試合を洗い出す（読み取り専用）。

保存済みのRAW（all_games_{日付}.xlsx / daily_pitch_data_{日付}.xlsx）を調べ、次のような試合を列挙する。
  状態が「試合終了」でない  : 試合中（9回裏・延長など）に取得して、そのまま止まっている
  スコアボード無し          : 試合情報だけ取れて成績が取れていない（中止・ノーゲームは除く）
  投手/打撃成績無し         : 同上
  投球データ無し            : daily_pitch_data にその試合が1球も無い（日ごと全滅も含む）
  投球データ最終回不一致    : 投球データの最終イニングがスコアボードの回数と違う（途中で切れている）
  投球データ不足 / 過多     : 投手成績の投球数の合計と、投球データの行数が大きく違う
                            （行数が数球多いのは通常運用でも出るため、不足は4球超・過多は8%超で判定）
  スコア不一致              : 基本情報の得点とスコアボードの合計が合わない
  日程にあるのにRAWに無い   : --schedule 指定時。Yahoo!の日程ページにある試合IDがRAWに1つも無い

使い方（GitHub Actions では data/ にデータ側リポジトリが展開される）:
  python baseball/scripts/find_incomplete_npb_games.py --year 2026
  python baseball/scripts/find_incomplete_npb_games.py --year 2026 --schedule     # 日程との突き合わせも行う（要ネット）
  python baseball/scripts/find_incomplete_npb_games.py --year 2026 --out report.json

出力の最後に「再取得が必要な 日付×1軍/2軍」の一覧を出す。再取得は
  python baseball/scripts/run.py --date 2026-07-02 --steps games pitch datamart --1軍
のように日付単位で行う（games ステップはその日の全試合を取り直す）。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd

_SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DATA_DIR = _SCRIPT_DIR.parent.parent / "data" / "baseball" / "プロ野球"
LEVELS = (("ichi", "1軍"), ("ni", "2軍"))

PITCH_SHORT_ABS = 4        # 成績の投球数 − 投球データ行数 がこれを超えたら「不足」
PITCH_OVER_RATIO = 0.08    # 投球データ行数が成績の投球数を8%超えたら「過多」


def _outs(ip) -> int:
    try:
        w, _, f = str(ip).partition(".")
        return int(w or 0) * 3 + int(f or 0)
    except Exception:  # noqa: BLE001
        return 0


def _num(v):
    n = pd.to_numeric(re.sub(r"[Xx×]", "", str(v)), errors="coerce")
    return None if pd.isna(n) else int(n)


def _read(path: Path, sheet=None):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001
        return None


def scan_date(raw_dir: Path, date: str) -> list[dict]:
    """1日ぶんのRAWを調べて、試合ごとの結果（problems付き）を返す。"""
    ag_path = raw_dir / f"all_games_{date}.xlsx"
    if not ag_path.exists():
        return [{"date": date, "gid": None, "problems": ["all_games無し"]}]
    x = _read(ag_path)
    if not isinstance(x, dict):
        return [{"date": date, "gid": None, "problems": ["all_gamesが読めない"]}]
    info = x.get("試合基本情報", pd.DataFrame())
    sb_all = x.get("スコアボード", pd.DataFrame())
    pit_all = x.get("投手成績", pd.DataFrame())
    bat_all = x.get("打撃成績", pd.DataFrame())
    pitch_path = raw_dir / f"daily_pitch_data_{date}.xlsx"
    pitch_df = _read(pitch_path, sheet=0) if pitch_path.exists() else None

    out = []
    for _, g in info.iterrows():
        gid = g["試合ID"]
        st = str(g.get("試合状態", ""))
        rec = {"date": date, "gid": str(int(gid)) if pd.notna(gid) else None,
               "home": g.get("ホームチーム"), "away": g.get("アウェイチーム"), "status": st, "problems": []}
        if "中止" in st or "ノーゲーム" in st:
            rec["cancelled"] = True
            out.append(rec)
            continue
        if st != "試合終了":
            rec["problems"].append(f"状態={st}")
        sb = sb_all[sb_all["試合ID"] == gid] if len(sb_all) else sb_all
        innings = None
        if len(sb) < 2:
            rec["problems"].append("スコアボード無し")
        else:
            cols = [c for c in sb.columns if re.fullmatch(r"\d+回", str(c))]
            innings = sum(1 for c in cols if sb[c].notna().any())
            for _, r in sb.iterrows():
                tot = sum(_num(r[c]) or 0 for c in cols)
                tt = pd.to_numeric(r.get("計"), errors="coerce")
                if pd.notna(tt) and tot != int(tt):
                    rec["problems"].append(f"回別合計≠計({r['チーム']}:{tot}/{int(tt)})")
            try:
                hs = int(pd.to_numeric(g["ホーム得点"]))
                if hs not in (int(pd.to_numeric(sb.iloc[0]["計"])), int(pd.to_numeric(sb.iloc[1]["計"]))):
                    rec["problems"].append("スコア不一致(基本情報)")
            except Exception:  # noqa: BLE001
                pass
        p = pit_all[pit_all["試合ID"] == gid] if len(pit_all) else pit_all
        b = bat_all[bat_all["試合ID"] == gid] if len(bat_all) else bat_all
        if len(p) == 0:
            rec["problems"].append("投手成績無し")
        if len(b) == 0:
            rec["problems"].append("打撃成績無し")
        stat_pitches = int(pd.to_numeric(p["投球数"], errors="coerce").fillna(0).sum()) if len(p) else None
        if pitch_df is None:
            rec["problems"].append("投球データ無し")
        else:
            q = pitch_df[pitch_df["試合ID"] == gid]
            rec["pitch_rows"] = len(q)
            if len(q) == 0:
                rec["problems"].append("投球データ無し")
            else:
                if stat_pitches:
                    diff = stat_pitches - len(q)
                    if diff > PITCH_SHORT_ABS:
                        rec["problems"].append(f"投球データ不足(成績{stat_pitches}/データ{len(q)})")
                    if -diff > PITCH_OVER_RATIO * stat_pitches:
                        rec["problems"].append(f"投球データ過多(成績{stat_pitches}/データ{len(q)})")
                last = q.iloc[-1]
                m = re.match(r"(\d+)", str(last["イニング"]))
                if m and innings and int(m.group(1)) != innings:
                    rec["problems"].append(f"投球データ最終回不一致({last['イニング']}{last['表/裏']}/スコアボード{innings}回)")
        out.append(rec)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", default="2026")
    ap.add_argument("--base", default=str(BASE_DATA_DIR), help="RAWのあるベースフォルダ（既定: data/baseball/プロ野球）")
    ap.add_argument("--schedule", action="store_true", help="Yahoo!の日程ページと突き合わせ、RAWに無い試合も探す")
    ap.add_argument("--out", default="", help="結果をJSONで保存するパス")
    args = ap.parse_args()

    results = []
    raw_ids_by_key: dict = {}
    for lg, lv in LEVELS:
        lv_dir = Path(args.base) / f"{args.year}年" / lv
        if not lv_dir.is_dir():
            continue
        for type_dir in sorted(p for p in lv_dir.iterdir() if p.is_dir()):
            raw_root = type_dir / "raw"
            if not raw_root.is_dir():
                continue
            for dpath in sorted(p for p in raw_root.iterdir() if p.is_dir()):
                recs = scan_date(dpath, dpath.name)
                for r in recs:
                    r["level"], r["type"] = lv, type_dir.name
                    if r.get("gid"):
                        raw_ids_by_key.setdefault((lv, dpath.name), set()).add(r["gid"])
                results += recs
    print(f"調査した試合: {len([r for r in results if r.get('gid')])}件 / 中止: {len([r for r in results if r.get('cancelled')])}件")

    missing_games = []
    if args.schedule:
        sys.path.insert(0, str(_SCRIPT_DIR))
        import run as npb  # noqa: WPS433  （スクレイパー本体。日程ページの取得関数を使う）
        for (lv, date) in sorted(raw_ids_by_key):
            lg = "ichi" if lv == "1軍" else "ni"
            try:
                npb.set_league_dirs(lg, date)
                ids = [str(i) for i in npb.get_game_ids()]
            except Exception as e:  # noqa: BLE001
                print(f"[WARN] 日程取得失敗 {lv} {date}: {e}")
                continue
            # 試合種別（レギュラー/交流戦…）が複数フォルダに分かれていても、同じ日の全フォルダのIDで比較する
            have = set(raw_ids_by_key.get((lv, date), set()))
            for gid in ids:
                if gid not in have:
                    missing_games.append({"date": date, "level": lv, "gid": gid, "problems": ["日程にあるのにRAWに無い"]})

    bad = [r for r in results if r.get("problems") and not r.get("cancelled")] + missing_games
    print(f"\n=== 問題のある試合: {len(bad)}件 ===")
    for r in bad:
        print(f"{r['level']} {r['date']} {r.get('gid')} {r.get('home','')}-{r.get('away','')} :: {' / '.join(r['problems'])}")

    need = sorted({(r["level"], r["date"]) for r in bad})
    print(f"\n=== 再取得が必要な 日付×リーグ: {len(need)}件 ===")
    for lv, d in need:
        flag = "--1軍" if lv == "1軍" else "--2軍"
        print(f"python baseball/scripts/run.py --date {d} --steps games pitch datamart {flag}")
    if args.out:
        Path(args.out).write_text(json.dumps({"problems": bad, "refetch": need}, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
