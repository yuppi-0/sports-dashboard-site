"""npb.jp の個人成績ページの解析とカード生成のテスト（通信なし）。"""
import sys
from pathlib import Path

from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "baseball" / "scripts"))
import npbjp_history as nh  # noqa: E402

BAT_2025 = """<table><tr><th>選手</th><th>試合</th><th>打席</th><th>打数</th><th>得点</th><th>安打</th><th>二塁打</th><th>三塁打</th><th>本塁打</th>
<th>塁打</th><th>打点</th><th>盗塁</th><th>盗塁刺</th><th>犠打</th><th>犠飛</th><th>四球</th><th>故意四</th><th>死球</th><th>三振</th><th>併殺打</th><th>打率</th><th>長打率</th><th>出塁率</th></tr>
<tr><td>*秋山　翔吾</td><td>64</td><td>157</td><td>145</td><td>12</td><td>38</td><td>6</td><td>0</td><td>1</td><td>47</td><td>5</td><td>0</td><td>0</td><td>1</td><td>0</td><td>9</td><td>0</td><td>2</td><td>27</td><td>2</td><td>.262</td><td>.324</td><td>.314</td></tr>
<tr><td>會澤　翼</td><td>24</td><td>55</td><td>43</td><td>4</td><td>5</td><td>1</td><td>0</td><td>0</td><td>6</td><td>1</td><td>0</td><td>1</td><td>3</td><td>0</td><td>7</td><td>2</td><td>2</td><td>16</td><td>1</td><td>.116</td><td>.140</td><td>.269</td></tr></table>"""
# 2021年は先頭に左右打の記号の列があり、見出しが「選　手」（全角空白入り）
BAT_2021 = """<table><tr><td>* 左打　+ 左右打</td></tr></table><table><tr><th></th><th>選　手</th><th>試合</th><th>打席</th><th>打数</th><th>得点</th><th>安打</th><th>二塁打</th><th>三塁打</th><th>本塁打</th>
<th>塁打</th><th>打点</th><th>盗塁</th><th>盗塁刺</th><th>犠打</th><th>犠飛</th><th>四球</th><th>故意四</th><th>死球</th><th>三振</th><th>併殺打</th><th>打率</th><th>長打率</th><th>出塁率</th></tr>
<tr><td></td><td>會澤　翼</td><td>70</td><td>204</td><td>180</td><td>18</td><td>46</td><td>10</td><td>0</td><td>3</td><td>65</td><td>22</td><td>0</td><td>0</td><td>1</td><td>3</td><td>16</td><td>2</td><td>4</td><td>33</td><td>8</td><td>.256</td><td>.361</td><td>.325</td></tr>
<tr><td>*</td><td>安部　友裕</td><td>85</td><td>170</td><td>151</td><td>8</td><td>38</td><td>5</td><td>0</td><td>1</td><td>46</td><td>12</td><td>2</td><td>1</td><td>1</td><td>2</td><td>15</td><td>2</td><td>1</td><td>47</td><td>2</td><td>.252</td><td>.305</td><td>.320</td></tr></table>"""
PIT = """<table><tr><th>選手</th><th>登板</th><th>勝利</th><th>敗北</th><th>セーブ</th><th>ホールド</th><th>ＨＰ</th><th>完投</th><th>完封勝</th><th>無四球</th><th>勝率</th><th>打者</th><th>投球回</th><th>安打</th><th>本塁打</th><th>四球</th><th>故意四</th><th>死球</th><th>三振</th><th>暴投</th><th>ボーク</th><th>失点</th><th>自責点</th><th>防御率</th></tr>
<tr><td>大瀬良　大地</td><td>23</td><td>7</td><td>9</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>.438</td><td>571</td><td>134.2</td><td>131</td><td>11</td><td>38</td><td>5</td><td>2</td><td>88</td><td>4</td><td>0</td><td>57</td><td>52</td><td>3.48</td></tr>
<tr><td>アドゥワ　誠</td><td>4</td><td>0</td><td>3</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>.000</td><td>88</td><td>19</td><td>26</td><td>2</td><td>5</td><td>0</td><td>1</td><td>11</td><td>0</td><td>0</td><td>14</td><td>11</td><td>5.21</td></tr></table>"""


def soup(html):
    return BeautifulSoup(html, "html.parser")


def test_parse_batting_both_layouts():
    for html in (BAT_2025, BAT_2021):
        rows = nh.parse_table(soup(html), "bat")
        assert len(rows) == 2 and all(r["選手"] in ("秋山 翔吾", "會澤 翼", "安部 友裕") for r in rows)
    r = nh.parse_table(soup(BAT_2025), "bat")[0]
    assert (r["選手"], r["試合"], r["安打"], r["打率"]) == ("秋山 翔吾", "64", "38", ".262")


def test_merge_batters_computes_official_rates_and_ops():
    rows = [dict(r, team="広島") for r in nh.parse_table(soup(BAT_2025), "bat")]
    b = nh.merge_batters(rows)["秋山 翔吾"]
    assert (b["avg"], b["obp"], b["slg"]) == (.262, .314, .324)          # npb.jp 公式の表と一致
    assert b["ops"] == round(49 / 156 + 47 / 145, 3)                      # 丸める前の合算
    c = nh.merge_batters(rows)["會澤 翼"]
    assert c["obp"] == .269                                              # 故意四球は四球に含まれる


def test_merge_traded_player_sums_both_teams():
    r = nh.parse_table(soup(BAT_2025), "bat")[0]
    rows = [dict(r, team="広島"), dict(r, team="阪神", 試合="10")]
    d = nh.merge_batters(rows)["秋山 翔吾"]
    assert d["試合"] == 74 and d["安打"] == 76 and d["team"] == "広島"


def test_pitchers_and_cards():
    prows = [dict(r, team="広島") for r in nh.parse_table(soup(PIT), "pit")]
    p = nh.merge_pitchers(prows)
    o = p["大瀬良 大地"]
    assert (o["innings"], o["era"], o["role"]) == ("134.2", 3.48, "先発")
    assert p["アドゥワ 誠"]["role"] == "先発"        # 4登板19回＝1登板4.7回
    reliever = nh.merge_pitchers([dict(prows[0], 選手="中継 太郎", 登板="50", 投球回="48.0")])["中継 太郎"]
    assert reliever["role"] == "中継ぎ"
    pi, pc = nh.build_pitcher_cards(p, 2025)
    assert {x["id"] for x in pi} == {"大瀬良_大地", "アドゥワ_誠"}
    assert pc["大瀬良_大地"]["k"] == 88 and pc["大瀬良_大地"]["source"] == "npb.jp"
    # 投手は打席が少なければ打者カードを作らない（アドゥワは打席3）
    brows = [dict(r, team="広島") for r in nh.parse_table(soup(BAT_2025), "bat")]
    bi, bc = nh.build_batter_cards(nh.merge_batters(brows), p, 2025)
    assert {x["id"] for x in bi} == {"秋山_翔吾", "會澤_翼"}
    assert bc["秋山_翔吾"]["seasons"]["2025"]["overall"]["pa"] == 157
