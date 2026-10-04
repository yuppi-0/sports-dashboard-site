"""NPB取得まわりの再発防止テスト（通信なし）。

  pip install -r requirements.txt pytest
  pytest -q tests
"""
import datetime
import sys
from pathlib import Path

import pandas as pd
import pytest
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import run as npb  # noqa: E402
import find_incomplete_npb_games as chk  # noqa: E402


def soup(html: str):
    return BeautifulSoup(html, "html.parser")


# ── 投球データの重複除去 ──
def test_dedupe_pitch_rows_keeps_first_and_distinguishes_pitchers():
    rows = [
        {"試合ID": "1", "投手名": "a", "通算投球数": "1", "x": 1},
        {"試合ID": "1", "投手名": "a", "通算投球数": "2", "x": 2},
        {"試合ID": "1", "投手名": "a", "通算投球数": "1", "x": 3},   # 同じ打席の別ページで再掲された同じ投球
        {"試合ID": "1", "投手名": "b", "通算投球数": "1", "x": 4},   # 別の投手の1球目は残す
        {"試合ID": "2", "投手名": "a", "通算投球数": "1", "x": 5},   # 別の試合は残す
    ]
    assert [r["x"] for r in npb.dedupe_pitch_rows(rows)] == [1, 2, 4, 5]


# ── 試合の無い日に直近の試合日が表示される問題 ──
SCHEDULE_HTML = """<html><body><h2 class="bb-head01__title">9月27日（日）</h2>
<div id="gm_card"><a href="/npb/game/2021041146/score">a</a><a href="/npb/game/2021041147/score">b</a></div></body></html>"""


def test_schedule_page_for_other_day_is_treated_as_no_games(monkeypatch):
    monkeypatch.setattr(npb, "get_soup", lambda url, *a, **k: soup(SCHEDULE_HTML))
    ok, ids = npb._fetch_schedule_ids("https://x/schedule/farm/official?date=2026-10-01")
    assert ok and ids == []


def test_schedule_page_for_same_day_returns_games(monkeypatch):
    monkeypatch.setattr(npb, "get_soup", lambda url, *a, **k: soup(SCHEDULE_HTML))
    ok, ids = npb._fetch_schedule_ids("https://x/schedule/farm/official?date=2026-09-27")
    assert ok and ids == ["2021041146", "2021041147"]


def test_game_page_date_and_wrong_date_game_is_skipped(monkeypatch):
    page = soup("<html><head><title>2026年9月27日 東北楽天ゴールデンイーグルスvs.東京ヤクルトスワローズ - スポーツナビ</title></head></html>")
    assert npb.game_page_date(page) == "2026-09-27"
    assert npb.game_page_date(soup("<html></html>")) is None
    monkeypatch.setattr(npb, "get_soup", lambda url, *a, **k: page)
    npb.WRONG_DATE_GAME_IDS.clear()
    assert npb.scrape_game_data("2021041146", expected_date="2026-10-01") is None
    assert "2021041146" in npb.WRONG_DATE_GAME_IDS


# ── チェック（find_incomplete_npb_games） ──
@pytest.mark.parametrize("v,expected", [(0, True), (3, True), (0.0, True), ("4", True), ("0X", False), ("2x", False), ("-", False), (float("nan"), False)])
def test_played(v, expected):
    assert chk._played(v) is expected


def _write_game(tmp_path, date, gid, pitch_rows, sb_rows):
    raw = tmp_path / date
    raw.mkdir(parents=True)
    info = pd.DataFrame([{"試合ID": gid, "ホームチーム": "H", "アウェイチーム": "A", "試合状態": "試合終了", "ホーム得点": 1}])
    pit = pd.DataFrame([{"試合ID": gid, "投手": "p", "投球数": len(pitch_rows)}])
    bat = pd.DataFrame([{"試合ID": gid, "打者": "b"}])
    with pd.ExcelWriter(raw / f"all_games_{date}.xlsx") as w:
        info.to_excel(w, sheet_name="試合基本情報", index=False)
        pd.DataFrame(sb_rows).to_excel(w, sheet_name="スコアボード", index=False)
        pit.to_excel(w, sheet_name="投手成績", index=False)
        bat.to_excel(w, sheet_name="打撃成績", index=False)
    pd.DataFrame(pitch_rows).to_excel(raw / f"daily_pitch_data_{date}.xlsx", index=False)
    return raw


def _pitch(gid, inning, tb, n, pitcher="p"):
    return {"試合ID": gid, "イニング": f"{inning}回", "表/裏": tb, "投手名": pitcher, "打者名": "b", "打席内球数": n, "通算投球数": n}


def test_scan_flags_duplicate_pitch_rows(tmp_path):
    sb = [{"試合ID": 1, "チーム": "A", "1回": 0, "計": 0}, {"試合ID": 1, "チーム": "H", "1回": 1, "計": 1}]
    pitches = [_pitch(1, 1, "裏", 1), _pitch(1, 1, "裏", 2), _pitch(1, 1, "裏", 2)]   # 2球目が重複
    raw = _write_game(tmp_path, "2026-04-01", 1, pitches, sb)
    recs = chk.scan_date(raw, "2026-04-01")
    assert any("投球データ重複" in p for r in recs for p in r["problems"])


def test_scan_does_not_flag_called_game_with_0X(tmp_path):
    # 5回裏で終了（コールド）。スコアボードの6列目は「0X」＝行われなかった裏
    sb = [{"試合ID": 1, "チーム": "A", **{f"{i}回": 0 for i in range(1, 6)}, "6回": None, "計": 0},
          {"試合ID": 1, "チーム": "H", **{f"{i}回": 0 for i in range(1, 5)}, "5回": 1, "6回": "0X", "計": 1}]
    pitches = [_pitch(1, 5, "裏", 1)]
    raw = _write_game(tmp_path, "2026-04-02", 1, pitches, sb)
    recs = chk.scan_date(raw, "2026-04-02")
    assert not [p for r in recs for p in r["problems"] if "最終回不一致" in p]


# ── 別の日付と同じ試合を保存していないか ──
def test_cross_date_problems_flags_only_the_later_dates():
    ids = {"2026-09-27": {"1", "2", "3"}, "2026-09-28": {"1", "2", "3"}, "2026-10-01": {"1", "2", "3"}, "2026-09-26": {"9"}}
    got = chk.cross_date_problems(ids, None)
    assert [g[0] for g in got] == ["2026-09-28", "2026-10-01"]
    scoped = chk.cross_date_problems(ids, (datetime.date(2026, 10, 1), datetime.date(2026, 10, 1)))
    assert [g[0] for g in scoped] == ["2026-10-01"]


def test_explicit_game_ids_only_for_saved_games_of_the_date(monkeypatch):
    """試合IDを明示した再取得でも、その日に保存済みの試合だけ投球データを取る（日付範囲＋試合ID指定で全日付に保存されていた）。"""
    import run as npb
    monkeypatch.setattr(npb, "known_game_folders", lambda *a, **k: {"2021044743": "レギュラーシーズン"})
    monkeypatch.setattr(npb, "schedule_folder_map", lambda *a, **k: {})
    seen = []
    monkeypatch.setattr(npb, "_run_pitch_scraper_impl", lambda gids, retry: seen.append(list(gids)) or "p")
    npb.run_pitch_scraper(target_game_ids=["2021044743", "2021039035"])
    assert seen == [["2021044743"]]
    seen.clear()
    monkeypatch.setattr(npb, "known_game_folders", lambda *a, **k: {})
    assert npb.run_pitch_scraper(target_game_ids=["2021039035"]) == ""
    assert seen == []


# ── 取りこぼし取り直しの打ち切り（実行時間を抑える） ──
def test_backfill_tries_each_game_only_once():
    import backfill_missing_npb_games as bf
    now = 1_000_000.0
    assert bf.eligible({}, "2軍", "2026-09-01", now)                                    # 初めては取り直す
    st = {"2軍|2026-09-01": {"tries": 1, "last": now - 99 * 3600}}
    assert not bf.eligible(st, "2軍", "2026-09-01", now)                                # 1回試したら、時間が経っても再試行しない


def test_backfill_state_roundtrip_and_raw_exists(tmp_path):
    import backfill_missing_npb_games as bf
    sp = tmp_path / "_backfill_state.json"
    assert bf.load_state(sp) == {}
    bf.save_state(sp, {"1軍|2026-09-01": {"tries": 1, "last": 1.0}})
    assert bf.load_state(sp)["1軍|2026-09-01"]["tries"] == 1
    d = tmp_path / "2026年" / "2軍" / "公式戦" / "raw" / "2026-09-02"
    d.mkdir(parents=True)
    assert not bf.raw_exists(tmp_path, 2026, "2軍", "2026-09-02")
    (d / "all_games_2026-09-02.xlsx").write_bytes(b"x")
    assert bf.raw_exists(tmp_path, 2026, "2軍", "2026-09-02")


def test_backfill_refetches_only_the_missing_games(monkeypatch):
    import backfill_missing_npb_games as bf
    calls = []
    monkeypatch.setattr(bf.subprocess, "run", lambda cmd, **kw: calls.append(cmd))
    bf.run_one("2軍", "2026-09-30", 100, ["2026040123", "2026040124"])
    assert calls[0][-3:] == ["2026040123", "2026040124", "--2軍"]          # 試合IDを渡す＝その試合だけ取得
    bf.run_one("1軍", "2026-09-30", 100)
    assert calls[1][-1] == "--1軍" and "2026040123" not in calls[1]          # IDが無ければ従来どおり日単位


def test_backfill_checks_only_previous_day_not_today():
    import datetime
    import backfill_missing_npb_games as bf
    # 当日（対象日）は含めず、前日だけを見る（当日分は翌日の実行が前日として確認する）
    assert bf.check_range("2026-10-04", 1) == (datetime.date(2026, 10, 3), datetime.date(2026, 10, 3))
    assert bf.check_range("2026-10-04", 3) == (datetime.date(2026, 10, 1), datetime.date(2026, 10, 3))   # 広く調べるとき


def test_backfill_groups_every_problem_game_by_id():
    import backfill_missing_npb_games as bf
    problems = [
        {"level": "1軍", "date": "2026-10-03", "gid": "2026040001", "problems": ["投球データ最終回不一致(8回表/スコアボード9回)"]},
        {"level": "1軍", "date": "2026-10-03", "gid": "2026040002", "problems": ["日程にあるのにRAWに無い"]},
        {"level": "1軍", "date": "2026-10-03", "gid": "2026040001", "problems": ["投球データ不足(成績120/データ100)"]},     # 同じ試合は1つにまとめる
        {"level": "2軍", "date": "2026-10-03", "gid": None, "problems": ["all_games無し"]},                              # その日が丸ごと無い
        {"level": "2軍", "date": "2026-10-02", "gid": None, "problems": ["別の日付(2026-10-01)と同じ試合を保存（3試合中3試合）"]},   # 試合を特定できない＝取り直さない
        {"level": "1軍", "date": "2026-10-03", "gid": "2026040003", "cancelled": True, "problems": []},                  # 中止は対象外
    ]
    got = bf.group_problems(problems)
    assert [(lv, d, g) for lv, d, g, _ in got] == [("1軍", "2026-10-03", ["2026040001", "2026040002"]), ("2軍", "2026-10-03", [])]
