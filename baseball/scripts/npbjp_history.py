"""過去シーズン（2021〜）のNPB1軍の打者・投手のシーズン成績を npb.jp（NPB公式）から取得し、既存の画面が読むカードJSONを作る。

  python baseball/scripts/npbjp_history.py --years 2025 2024 --docs docs --raw data/baseball/プロ野球

Yahoo!は現在のシーズンしか表示できないため、過去シーズンは npb.jp の「個人打撃成績」「個人投手成績」（球団別）を使う。
取れるのは打撃・投球の成績（試合・打席・打数・安打・本塁打・打点・四死球・三振・打率・出塁率・長打率、投手は投球回・防御率・奪三振 等）で、
投球データ（球種・コース・空振り率など）と試合ごとの記録は無いので、カードのその項目は空になる。
出力:
  {raw}/{年}年/1軍/レギュラーシーズン/npbjp/season_{年}.json   取得した生データ
  {docs}/baseball/data/プロ野球/{年}年/1軍/レギュラーシーズン/{batter,pitcher}_cards_numeric/（index.json と選手別.json.gz）
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parent))
from jsonio import write_json  # noqa: E402

BASE = "https://npb.jp/bis"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; sports-dashboard)"}
# npb.jp のファイル名の球団コード → 当サイトの球団名
TEAMS = [("g", "巨人"), ("t", "阪神"), ("db", "DeNA"), ("c", "広島"), ("s", "ヤクルト"), ("d", "中日"),
         ("h", "ソフトバンク"), ("f", "日本ハム"), ("m", "ロッテ"), ("e", "楽天"), ("b", "オリックス"), ("l", "西武")]
# 見出しの行が見つからないページ用の既定の列順（2021〜2025年のページで共通）
BAT_HEADS = ["選手", "試合", "打席", "打数", "得点", "安打", "二塁打", "三塁打", "本塁打", "塁打", "打点", "盗塁", "盗塁刺", "犠打", "犠飛",
             "四球", "故意四", "死球", "三振", "併殺打", "打率", "長打率", "出塁率"]
PIT_HEADS = ["選手", "登板", "勝利", "敗北", "セーブ", "ホールド", "ＨＰ", "完投", "完封勝", "無四球", "勝率", "打者", "投球回", "安打", "本塁打",
             "四球", "故意四", "死球", "三振", "暴投", "ボーク", "失点", "自責点", "防御率"]
BAT_INT = ["試合", "打席", "打数", "得点", "安打", "二塁打", "三塁打", "本塁打", "塁打", "打点", "盗塁", "盗塁刺", "犠打", "犠飛", "四球", "故意四", "死球", "三振", "併殺打"]
PIT_INT = ["登板", "勝利", "敗北", "セーブ", "ホールド", "ＨＰ", "完投", "完封勝", "無四球", "打者", "安打", "本塁打", "四球", "故意四", "死球", "三振", "暴投", "ボーク", "失点", "自責点"]
MIN_PA_TWO_WAY = 100     # 投手でもこの打席数以上なら打者カードも作る（二刀流）


def get_html(url: str, retries: int = 4):
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return BeautifulSoup(r.content, "html.parser", from_encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            if attempt == retries - 1:
                print(f"  [WARN] 取得失敗 {url}: {e}")
                return None
            time.sleep(2 ** attempt)
    return None


def _norm(s: str) -> str:
    return re.sub(r"[\s　]", "", s or "")


def clean_name(s: str) -> str:
    """'*秋山　翔吾' → '秋山 翔吾'（左右打の記号を除き、全角空白を半角1つにする）。"""
    s = re.sub(r"^[*＊+＋\s　]+", "", (s or "").strip())
    return re.sub(r"[\s　]+", " ", s).strip()


def _num(s: str):
    s = (s or "").strip().replace(",", "")
    if s in ("", "-", "－", "---"):
        return None
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def parse_table(soup, kind: str) -> list:
    """個人成績ページから [{見出し: 値}] を返す。kind='bat'（打率の列がある表）/'pit'（防御率の列がある表）。

    ページの形は年度で違う:
      ・2025年: 見出しの先頭が「選手」、行の先頭が選手名。投球回は1つのセル（<span>で整数部と小数部）
      ・2021年: 見出しの先頭に左右打の記号の空の列があり、行も先頭に記号の列
      ・2024年の投手: 行が <tr class="ststats"> で、選手名は <td class="stplayer">。投球回は整数部と小数部の2つのセルに分かれる
    見出し・名前の位置から値のセルを数え直して、どの形でも同じ辞書にする。"""
    key = "打率" if kind == "bat" else "防御率"
    best = None
    for t in soup.find_all("table"):
        rows = t.find_all("tr")
        for i, tr in enumerate(rows):
            cells = [_norm(c.get_text()) for c in tr.find_all(["th", "td"])]
            if key in cells:
                if best is None or len(rows) > best[2]:
                    best = (cells, rows[i + 1:], len(rows))
                break
    if best:
        heads, rows, _ = best
        name_i = heads.index("選手") if "選手" in heads else 0
    else:
        # 見出しの行を読めないページ（2024年以前の投手成績など）は、データ行（tr.ststats）と既定の列順で読む
        rows = soup.find_all("tr", class_="ststats")
        heads, name_i = (BAT_HEADS if kind == "bat" else PIT_HEADS), 0
        if not rows:
            return []
    heads_after = heads[name_i + 1:]
    out = []
    for tr in rows:
        tds = tr.find_all(["th", "td"])
        texts = [c.get_text(strip=True) for c in tds]
        named = tr.find("td", class_="stplayer")
        if named is not None:
            pos = tds.index(named)
        else:
            pos = name_i
        if len(texts) <= pos + 1:
            continue
        name = clean_name(texts[pos])
        vals = texts[pos + 1:]
        if "投球回" in heads_after and len(vals) == len(heads_after) + 1:
            i = heads_after.index("投球回")           # 投球回が整数部と小数部の2セルに分かれているページ
            frac = vals[i + 1] if vals[i + 1].startswith(".") else ""
            vals = vals[:i] + [vals[i] + frac] + vals[i + 2:]
        if not name or len(vals) != len(heads_after):
            continue
        row = dict(zip(heads_after, vals))
        row["選手"] = name
        if row.get("試合" if kind == "bat" else "登板") not in (None, ""):
            out.append(row)
    return out


def fetch_year(year: int) -> dict:
    """球団別ページ（打撃・投手）を取得して {'batters': [...], 'pitchers': [...]} を返す（各行に team を付ける）。"""
    res = {"year": year, "batters": [], "pitchers": []}
    for code, team in TEAMS:
        for kind, page, outkey in (("bat", "idb1", "batters"), ("pit", "idp1", "pitchers")):
            soup = get_html(f"{BASE}/{year}/stats/{page}_{code}.html")
            time.sleep(0.4)
            if not soup:
                print(f"  [WARN] {year} {team} {outkey}: ページなし")
                continue
            rows = parse_table(soup, kind)
            if not rows:   # 解析できない形のページ：構造を調べられるよう見出しを出す
                heads = [[_norm(c.get_text()) for c in (t.find("tr") or t).find_all(["th", "td"])][:8] for t in soup.find_all("table")[:4]]
                print(f"  [WARN] {year} {team} {outkey}: 表を解析できません title={soup.title.get_text(strip=True) if soup.title else None} 表の見出し={heads}")
            for row in rows:
                row["team"] = team
                res[outkey].append(row)
    return res


def _ip_outs(ip: str) -> int:
    m = re.fullmatch(r"(\d+)(?:\.([012]))?", str(ip).strip())
    return int(m.group(1)) * 3 + int(m.group(2) or 0) if m else 0


def _ip_str(outs: int) -> str:
    return f"{outs // 3}.{outs % 3}"


def merge_batters(rows: list) -> dict:
    """同じ選手（移籍で複数球団に載る）を合算し、{名前: 成績dict} を返す。所属は試合数の多い球団。"""
    out: dict = {}
    for r in rows:
        name = r["選手"]
        d = out.setdefault(name, {"name": name, "teams": {}, **{k: 0 for k in BAT_INT}})
        for k in BAT_INT:
            d[k] += _num(r.get(k)) or 0
        d["teams"][r["team"]] = d["teams"].get(r["team"], 0) + (_num(r.get("試合")) or 0)
    for d in out.values():
        d["team"] = max(d["teams"].items(), key=lambda kv: kv[1])[0]
        ab, h, bb, hbp, sf = d["打数"], d["安打"], d["四球"], d["死球"], d["犠飛"]
        tb = d["塁打"]
        den = ab + bb + hbp + sf
        d["avg"] = round(h / ab, 3) if ab else None
        d["obp"] = round((h + bb + hbp) / den, 3) if den else None
        d["slg"] = round(tb / ab, 3) if ab else None
        d["ops"] = round((h + bb + hbp) / den + tb / ab, 3) if den and ab else None
    return out


def merge_pitchers(rows: list) -> dict:
    out: dict = {}
    for r in rows:
        name = r["選手"]
        d = out.setdefault(name, {"name": name, "teams": {}, "outs": 0, **{k: 0 for k in PIT_INT}})
        for k in PIT_INT:
            d[k] += _num(r.get(k)) or 0
        d["outs"] += _ip_outs(r.get("投球回"))
        d["teams"][r["team"]] = d["teams"].get(r["team"], 0) + (_num(r.get("登板")) or 0)
    for d in out.values():
        d["team"] = max(d["teams"].items(), key=lambda kv: kv[1])[0]
        d["innings"] = _ip_str(d["outs"])
        ip = d["outs"] / 3
        d["era"] = round(d["自責点"] * 9 / ip, 2) if ip else None
        d["whip"] = round((d["安打"] + d["四球"]) / ip, 2) if ip else None
        d["k9"] = round(d["三振"] * 9 / ip, 2) if ip else None
        # 先発・中継ぎは npb.jp の表に先発数が無いので、1登板あたりの投球回で判定する（3回以上で先発）
        d["role"] = "先発" if d["登板"] and ip / d["登板"] >= 3.0 else "中継ぎ"
    return out


def slug(name: str) -> str:
    s = re.sub(r"[^\w]+", "_", name.strip().lower())
    return s.strip("_") or "unknown"


def _pct(n, d):
    return round(n / d * 100, 1) if d else None


def build_batter_cards(batters: dict, pitchers: dict, year: int) -> tuple:
    """([index行], {id: カードJSON}) を返す。投手（投手成績に載る選手）は打席が少なければ打者カードを作らない。"""
    idx, cards = [], {}
    for name, b in sorted(batters.items()):
        if b["打席"] <= 0:
            continue
        if name in pitchers and b["打席"] < MIN_PA_TWO_WAY:
            continue
        pid = slug(name)
        k_pct, bb_pct = _pct(b["三振"], b["打席"]), _pct(b["四球"], b["打席"])
        overall = {"games": b["試合"], "pa": b["打席"], "ab": b["打数"], "h": b["安打"], "hr": b["本塁打"], "bb": b["四球"],
                   "k": b["三振"], "rbi": b["打点"], "sb": b["盗塁"], "cs": b["盗塁刺"], "avg": b["avg"], "obp": b["obp"],
                   "slg": b["slg"], "ops": b["ops"], "k_pct": k_pct, "bb_pct": bb_pct, "hr_pct": _pct(b["本塁打"], b["打席"])}
        cards[pid] = {
            "name": name, "team": b["team"], "pos": None, "years": [str(year)], "latestYear": str(year),
            **{k: overall[k] for k in ("games", "pa", "ab", "h", "hr", "bb", "k", "rbi", "sb", "avg", "obp", "slg", "ops")},
            "k_pct_season": k_pct, "bb_pct_season": bb_pct,
            "rankings": {}, "game_log": [], "categories": [],
            "seasons": {str(year): {"team": b["team"], "pos": None, "bats": None, "overall": overall, "splits": {}}},
            "source": "npb.jp", "d2": b["二塁打"], "d3": b["三塁打"], "hbp": b["死球"], "sf": b["犠飛"], "sh": b["犠打"],
        }
        idx.append({"id": pid, "name": name, "team": b["team"], "pos": None, "categories": [], "games": b["試合"], "pa": b["打席"],
                    "avg": b["avg"], "obp": b["obp"], "slg": b["slg"], "ops": b["ops"], "hr": b["本塁打"], "rbi": b["打点"],
                    "sb": b["盗塁"], "k_pct": k_pct, "bb_pct": bb_pct, "war": None})
    return idx, cards


def build_pitcher_cards(pitchers: dict, year: int) -> tuple:
    idx, cards = [], {}
    for name, p in sorted(pitchers.items()):
        if p["outs"] <= 0 and p["登板"] <= 0:
            continue
        pid = slug(name)
        tbf = p["打者"]
        k_pct, bb_pct = _pct(p["三振"], tbf), _pct(p["四球"], tbf)
        kbb = _pct(p["三振"] - p["四球"], tbf)
        cards[pid] = {
            "name": name, "team": p["team"], "role": p["role"], "throws": None, "games": p["登板"], "innings": p["innings"],
            "era": p["era"], "war": None, "k_bb_pct": kbb, "gb_pct": None, "k": p["三振"], "bb": p["四球"],
            "k_pct_season": k_pct, "bb_pct_season": bb_pct, "kbb_pct_season": kbb,
            "pitch_evaluations_numeric": [], "season_pitch_detail": {}, "season_totals": {}, "season_course_detail": {},
            "season_course_locs": {}, "game_log": [], "vs_batters": [], "rankings": {}, "rankings_by_hand": {}, "categories": [],
            "source": "npb.jp", "h": p["安打"], "hr": p["本塁打"], "hbp": p["死球"], "r": p["失点"], "er": p["自責点"], "tbf": tbf,
            "w": p["勝利"], "l": p["敗北"], "sv": p["セーブ"], "hld": p["ホールド"], "whip": p["whip"], "k9": p["k9"],
        }
        idx.append({"id": pid, "name": name, "team": p["team"], "role": p["role"], "innings": p["innings"],
                    "innings_num": round(p["outs"] / 3, 1), "categories": [], "era": p["era"], "war": None, "k_bb_pct": kbb,
                    "k_pct": k_pct, "bb_pct": bb_pct})
    return idx, cards


def write_cards(docs: Path, year: int, kind: str, idx: list, cards: dict) -> None:
    d = docs / "baseball/data/プロ野球" / f"{year}年" / "1軍" / "レギュラーシーズン" / f"{kind}_cards_numeric"
    d.mkdir(parents=True, exist_ok=True)
    for f in d.glob("*.json.gz"):      # 古いカードが残らないよう、その年のカードは作り直す
        f.unlink()
    for pid, card in cards.items():
        write_json(str(d / f"{pid}.json"), card)
    (d / "index.json").write_text(json.dumps({"players": idx}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  {year}年 {kind}: {len(idx)}人 → {d}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--years", nargs="+", type=int, default=[2025, 2024, 2023, 2022, 2021])
    ap.add_argument("--docs", default="docs")
    ap.add_argument("--raw", default="data/baseball/プロ野球")
    ap.add_argument("--from-raw", action="store_true", help="取得せず、保存済みの生データからカードだけ作り直す")
    args = ap.parse_args()
    for year in args.years:
        raw_path = Path(args.raw) / f"{year}年" / "1軍" / "レギュラーシーズン" / "npbjp" / f"season_{year}.json"
        if args.from_raw and raw_path.exists():
            data = json.loads(raw_path.read_text(encoding="utf-8"))
        else:
            print(f"{year}年を取得します")
            data = fetch_year(year)
            raw_path.parent.mkdir(parents=True, exist_ok=True)
            raw_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        print(f"  {year}年: 打者の行 {len(data['batters'])} / 投手の行 {len(data['pitchers'])}")
        if not data["batters"] or not data["pitchers"]:
            print(f"::warning::{year}年のデータが取得できませんでした")
            continue
        b, p = merge_batters(data["batters"]), merge_pitchers(data["pitchers"])
        bi, bc = build_batter_cards(b, p, year)
        pi, pc = build_pitcher_cards(p, year)
        write_cards(Path(args.docs), year, "batter", bi, bc)
        write_cards(Path(args.docs), year, "pitcher", pi, pc)


if __name__ == "__main__":
    main()
