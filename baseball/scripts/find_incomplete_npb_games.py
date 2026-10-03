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
import datetime
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


def _played(v) -> bool:
    """スコアボードのセルが「実際に攻撃した回」を表すか。数字だけのセル。
    「X」付き（例 0X）は、裏の攻撃が不要で行われなかった回（サヨナラ・コールド等）の表記なので数えない。"""
    s = str(v)
    if re.search(r"[Xx×]", s):
        return False
    return bool(pd.notna(pd.to_numeric(s, errors="coerce")))


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
            # コールドゲーム等で終わった回の次の列に「-」「x」だけが入ることがあるので、数字の入っている回だけを数える
            innings = sum(1 for c in cols if sb[c].notna().any())
            # 最後の列に「0X」（行われなかった裏）があるか。コールド等で、最後の投球が「N回裏」・スコアボードがN+1列、になる
            last_col = [c for c in cols if sb[c].notna().any()][-1] if innings else None
            last_has_x = bool(last_col is not None and sb[last_col].astype(str).str.contains(r"[Xx×]").any())
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
                # 同じ投球の重複（取得処理の不具合で起きた。データマートは除去済みだがRAWには残る）
                if all(c in q.columns for c in ("投手名", "通算投球数")):
                    dup = int(q.duplicated(subset=["投手名", "通算投球数"]).sum())
                    if dup:
                        rec["problems"].append(f"投球データ重複({dup}行)")
                if stat_pitches:
                    diff = stat_pitches - len(q)
                    if diff > PITCH_SHORT_ABS:
                        rec["problems"].append(f"投球データ不足(成績{stat_pitches}/データ{len(q)})")
                    if -diff > PITCH_OVER_RATIO * stat_pitches:
                        rec["problems"].append(f"投球データ過多(成績{stat_pitches}/データ{len(q)})")
                last = q.iloc[-1]
                m = re.match(r"(\d+)", str(last["イニング"]))
                ended_after_bottom = (int(m.group(1)) == innings - 1 and str(last["表/裏"]) == "裏" and last_has_x) if (m and innings) else False
                if m and innings and int(m.group(1)) != innings and not ended_after_bottom:
                    rec["problems"].append(f"投球データ最終回不一致({last['イニング']}{last['表/裏']}/スコアボード{innings}回)")
        out.append(rec)
    return out


def diagnose(r: dict, base: str) -> None:
    """問題のある1試合について、なぜそうなったかの手掛かり（重複行・投手別の差・最後の打席など）を出す。"""
    lv_dir = Path(base) / f"{r['year']}年" / r["level"]
    raw = next(iter(sorted(lv_dir.glob(f"*/raw/{r['date']}"))), None)
    if raw is None:
        print("   (raw無し)")
        return
    gid = int(r["gid"])
    pitch_path = raw / f"daily_pitch_data_{r['date']}.xlsx"
    df = _read(pitch_path, sheet=0) if pitch_path.exists() else None
    ag = _read(raw / f"all_games_{r['date']}.xlsx") or {}
    if df is None:
        print("   投球データ無し")
        return
    q = df[df["試合ID"] == gid].reset_index(drop=True)
    pit = ag.get("投手成績", pd.DataFrame())
    pit = pit[pit["試合ID"] == gid] if len(pit) else pit
    keys = [c for c in ["イニング", "表/裏", "打者名", "打席内球数", "通算投球数", "球種", "球速", "1球結果"] if c in q.columns]
    dups = int(q.duplicated(subset=keys).sum())
    print(f"   投球行{len(q)} / 完全重複行{dups}")
    name_col = next((c for c in pit.columns if c in ("投手", "選手", "選手名", "名前")), None)
    if name_col and "投手名" in q.columns:
        by = q.groupby("投手名").size()
        for _, pr in pit.iterrows():
            nm = str(pr[name_col]).replace(" ", "").replace("\u3000", "")
            cnt = next((int(v) for k, v in by.items() if str(k).replace(" ", "").replace("\u3000", "") == nm), 0)
            st = int(pd.to_numeric(pr["投球数"], errors="coerce") or 0)
            if abs(cnt - st) > 0:
                print(f"   投手 {pr[name_col]}: データ{cnt}球 / 成績{st}球 (差{cnt - st:+d})")
                if cnt < st:
                    seqs = pd.to_numeric(q[q["投手名"].astype(str).str.replace(" ", "").str.replace("\u3000", "") == nm]["通算投球数"], errors="coerce").dropna().astype(int)
                    miss = [i for i in range(1, (int(seqs.max()) if len(seqs) else 0) + 1) if i not in set(seqs)]
                    print(f"     通算投球数の最大={int(seqs.max()) if len(seqs) else None} / 欠番={miss[:20]}")
    else:
        print(f"   [投手成績の列] {list(pit.columns)[:12]}")
    if len(q):
        last = q.iloc[-1]
        print(f"   最後の投球: {last['イニング']}{last['表/裏']} 打者={last.get('打者名')} 球数={last.get('通算投球数')} 結果={last.get('1球結果')} / 打席結果={last.get('打席完了結果')}")
    # 打席内球数の巻き戻り（同じ打席のページを二重に取ったときに出る）
    seq = pd.to_numeric(q["通算投球数"], errors="coerce") if "通算投球数" in q.columns else None
    if seq is not None:
        back = int((seq.diff() < 0).sum())
        print(f"   通算投球数が巻き戻る箇所: {back}")
    sp = ag.get("スコアプレー詳細", pd.DataFrame())
    sp = sp[sp["試合ID"] == gid] if len(sp) else sp
    sb = ag.get("スコアボード", pd.DataFrame())
    sb = sb[sb["試合ID"] == gid] if len(sb) else sb
    print(f"   スコアボード行{len(sb)} / スコアプレー行{len(sp)}")
    # スコアボードの中身（コールド・途中終了の判定に使う）
    cols = [c for c in sb.columns if re.fullmatch(r"\d+回", str(c))] + [c for c in ("計",) if c in sb.columns]
    for _, srow in sb.iterrows():
        print("   SB", srow.get("チーム"), [srow[c] for c in cols])
    st = ag.get("試合基本情報", pd.DataFrame())
    st = st[st["試合ID"] == gid] if len(st) else st
    if len(st):
        print("   試合状態:", st.iloc[0].get("試合状態"), "/ 試合情報:", st.iloc[0].get("試合情報"))


def _date_scope(spec: str):
    """'2026-05-13' または '2026-05-13:2026-05-20' → (開始, 終了)。空なら None。"""
    if not spec:
        return None
    a, _, b = spec.partition(":")
    d0 = datetime.date.fromisoformat(a)
    return d0, datetime.date.fromisoformat(b) if b else d0


def cross_date_problems(ids_by_date: dict, scope) -> list:
    """同じ試合IDが別の日付にも保存されている日付（後の日付側）を返す。scope があればその範囲の日付だけ。"""
    out = []
    first_seen: dict = {}
    for d in sorted(ids_by_date):
        for gid in sorted(ids_by_date[d]):
            first_seen.setdefault(gid, d)
    for d in sorted(ids_by_date):
        dd = datetime.date.fromisoformat(d)
        if scope and not (scope[0] <= dd <= scope[1]):
            continue
        dup = {g: first_seen[g] for g in ids_by_date[d] if first_seen[g] < d}
        if dup:
            src = sorted(set(dup.values()))
            out.append((d, len(dup), len(ids_by_date[d]), src))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--year", default="2026")
    ap.add_argument("--base", default=str(BASE_DATA_DIR), help="RAWのあるベースフォルダ（既定: data/baseball/プロ野球）")
    ap.add_argument("--schedule", action="store_true", help="Yahoo!の日程ページと突き合わせ、RAWに無い試合も探す")
    ap.add_argument("--date", default="", help="この日（または A:B の範囲）だけを調べる。毎日の自動更新後チェック用")
    ap.add_argument("--gh-warning", action="store_true", help="問題をGitHub Actionsの警告（::warning::）としても出す")
    ap.add_argument("--diagnose", action="store_true", help="問題のある試合について原因の手掛かりを出す")
    ap.add_argument("--out", default="", help="結果をJSONで保存するパス")
    args = ap.parse_args()

    scope = _date_scope(args.date)
    window = None
    if scope:
        window = (scope[0] - datetime.timedelta(days=14), scope[1] + datetime.timedelta(days=14))
    results = []
    raw_ids_by_key: dict = {}
    ids_by_type: dict = {}   # (リーグ, 試合種別) -> {日付: 試合IDの集合}
    for lg, lv in LEVELS:
        lv_dir = Path(args.base) / f"{args.year}年" / lv
        if not lv_dir.is_dir():
            continue
        for type_dir in sorted(p for p in lv_dir.iterdir() if p.is_dir()):
            raw_root = type_dir / "raw"
            if not raw_root.is_dir():
                continue
            for dpath in sorted(p for p in raw_root.iterdir() if p.is_dir()):
                try:
                    dd = datetime.date.fromisoformat(dpath.name)
                except ValueError:
                    continue
                in_scope = scope is None or (scope[0] <= dd <= scope[1])
                if window and not (window[0] <= dd <= window[1]):
                    continue
                ids_here = set()
                try:
                    _info = pd.read_excel(dpath / f"all_games_{dpath.name}.xlsx", sheet_name="試合基本情報")
                    ids_here = {str(int(g)) for g in _info["試合ID"].dropna()}
                except Exception:  # noqa: BLE001
                    pass
                if ids_here:
                    ids_by_type.setdefault((lv, type_dir.name), {})[dpath.name] = ids_here
                if not in_scope:
                    continue
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
        # RAWの日付フォルダが1つも無い日（取得が丸ごと抜けた日・中止後の振替試合の日）も日程と突き合わせる。
        # 日付フォルダがある日だけを見ていると、フォルダごと無い日の試合は永久に検出できない。
        check_keys = set(raw_ids_by_key)
        for lv in {lv for (lv, _gt) in ids_by_type}:
            raw_dates = sorted(datetime.date.fromisoformat(d)
                               for (lv2, _gt), by_date in ids_by_type.items() if lv2 == lv for d in by_date)
            if not raw_dates:
                continue
            start, end = (window[0], scope[1]) if scope else (raw_dates[0], raw_dates[-1])
            start = max(start, raw_dates[0])
            end = min(end, datetime.date.today())
            d = start
            while d <= end:
                check_keys.add((lv, d.isoformat()))
                d += datetime.timedelta(days=1)
        for (lv, date) in sorted(check_keys):
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
    # 同じ試合が別の日付にも保存されていないか（試合の無い日に直近の試合日の試合が保存される不具合の検出）
    for (lv, gt), by_date in ids_by_type.items():
        for d, n, total, src in cross_date_problems(by_date, scope):
            bad.append({"date": d, "level": lv, "gid": None, "home": "", "away": "",
                        "problems": [f"別の日付({', '.join(src)})と同じ試合を保存（{total}試合中{n}試合）"]})
    print(f"\n=== 問題のある試合: {len(bad)}件 ===")
    for r in bad:
        print(f"{r['level']} {r['date']} {r.get('gid')} {r.get('home','')}-{r.get('away','')} :: {' / '.join(r['problems'])}")

    if args.diagnose:
        print("\n=== 原因の手掛かり ===")
        for r in bad:
            if not r.get("gid") or r.get("cancelled"):
                continue
            r["year"] = args.year
            print(f"{r['level']} {r['date']} {r['gid']} :: {' / '.join(r['problems'])}")
            try:
                diagnose(r, args.base)
            except Exception as e:  # noqa: BLE001
                print(f"   [diagnose失敗] {e}")

    if args.gh_warning:
        for r in bad[:20]:
            print(f"::warning::NPB取得チェック {r['level']} {r['date']} {r.get('gid') or ''} {' / '.join(r['problems'])}")
        if not bad:
            print("NPB取得チェック: 問題なし")

    need = sorted({(r["level"], r["date"]) for r in bad})
    print(f"\n=== 再取得が必要な 日付×リーグ: {len(need)}件 ===")
    for lv, d in need:
        flag = "--1軍" if lv == "1軍" else "--2軍"
        print(f"python baseball/scripts/run.py --date {d} --steps games pitch datamart {flag}")
    if args.out:
        Path(args.out).write_text(json.dumps({"problems": bad, "refetch": need}, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
