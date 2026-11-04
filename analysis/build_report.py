"""生成完整研究报告 HTML（output/港股红利与AH抉择.html）。

数字全部从 output/ 下各分析结果表读取；先依次运行：
  fed_cycle_hk_dividend.py, spread_and_flows.py, high_rate_regime.py, ccass_holdings.py, ah_premium.py, ah_scenarios.py
"""
import base64
import html
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"
DST = OUT / "港股红利与AH抉择.html"


def rd(name, **kw):
    return pd.read_csv(OUT / name, **kw)


def img(name, alt):
    b = base64.b64encode((OUT / name).read_bytes()).decode()
    return f'<figure><div class="plate"><img src="data:image/png;base64,{b}" alt="{html.escape(alt)}"></div><figcaption>{html.escape(alt)}</figcaption></figure>'


def pct(v, d=1, sign=False):
    if pd.isna(v):
        return "—"
    s = f"{v * 100:+.{d}f}%" if sign else f"{v * 100:.{d}f}%"
    return s


def pp(v, d=2):
    return "—" if pd.isna(v) else f"{v * 100:+.{d}f}pp"


def num(v, d=1):
    return "—" if pd.isna(v) else f"{v:.{d}f}"


def table(head, rows, cls="", note=None):
    h = "".join(f"<th>{html.escape(str(x))}</th>" for x in head)
    b = ""
    for r in rows:
        tds = []
        for i, x in enumerate(r):
            x = str(x)
            c = ""
            if i > 0 and (x.startswith("+") or x.startswith("-") or x.startswith("−")):
                c = ' class="pos"' if x.startswith("+") else ' class="neg"'
            tds.append(f"<td{c}>{x}</td>")
        b += "<tr>" + "".join(tds) + "</tr>"
    n = f'<p class="tnote">{note}</p>' if note else ""
    return f'<div class="tw {cls}"><table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table></div>{n}'


def chip(label):
    k = {"买H": "h", "偏H": "hl", "无差别": "n", "买A": "a"}.get(label.split("（")[0], "n")
    return f'<span class="chip {k}">{html.escape(label)}</span>'


# ---------------- 数据 ----------------
hr = rd("20_高利率vs非高利率_年化.csv")
hrA = hr[hr["口径"] == "A_美债10Y≥4%"].set_index("环境")
ep = rd("22_高利率区间明细_美债4%.csv")
hrp = rd("21_高利率x联储阶段_年化.csv")
hrpA = hrp[hrp["口径"] == "A_美债10Y≥4%"].set_index("联储阶段")
ev = rd("03_首次调整事件窗口_超额HSI.csv")
snap = rd("16_各视角利差快照.csv").set_index("时点")
q5 = rd("07_美债利差分组_未来12M超额.csv")
g = rd("31_分组持股比例_港股市值加权.csv", parse_dates=["日期"])
cc = rd("32_高利率vs非高利率_中介净买入.csv")
epi = rd("34_高利率区间_累计中介净买入.csv")
pan = rd("35_面板回归_持股变化对利差.csv")
b41 = rd("41_溢价分档_未来12M_H减A.csv")
rob = rd("48_稳健性_不重叠年度样本.csv")
dec = rd("44_当前AH个股决策表.csv")
idx = rd("45_指数层面_A红利vs港股通高股息.csv", index_col=0)
i46 = rd("46_AH指数分档_未来12M_港股高股息减A红利.csv")
fxd = rd("50_港币兑人民币12M变化分布.csv")
fxr = rd("51_汇率情景回归.csv")
sc = rd("52_汇率情景_个股预期H减A.csv")
c53 = rd("53_历史_汇率x溢价_H减A均值.csv")
c53n = rd("53b_历史_汇率x溢价_样本数.csv")
nar = rd("54_溢价档_收窄概率.csv")
hl = rd("58_溢价均值回归半衰期.csv")
lg = rd("57_收窄概率模型.csv")
cur = rd("59_当前个股_溢价收窄概率.csv")

A_hi, A_lo = hrA.loc["高利率"], hrA.loc["非高利率"]
gA = g[g["分组"] == "A"].set_index("日期")
gA_2017 = gA.loc["2017-12-29"]
gA_now = gA.iloc[-1]
ccA = cc[cc["研究分组"] == "A"].set_index("高利率")
ccB = cc[cc["研究分组"] == "B"].set_index("高利率")
eB = epi[(epi["分组"] == "B") & epi["区间"].str.startswith("2024-10")].iloc[0]
pan_hi = pan[(pan["因变量"] == "Δ非港股通_pp") & (pan["样本"] == "A+B 高息股") & (pan["变量"] == "高")].iloc[0]
pan_sb = pan[(pan["因变量"] == "Δ港股通_pp") & (pan["样本"] == "A+B 高息股") & (pan["变量"] == "高")].iloc[0]
now = snap.loc["当前"]
ii = idx.dropna(subset=["AH指数"]).iloc[-1]
rb = rob[(rob["因变量"] == "H减A_个人") & (rob["固定效应"] == "无")].iloc[0]
fx = dict(zip(fxd["分位"].astype(str), fxd["港币兑人民币12M变化"]))
fx15 = fxr[(fxr["样本"] == "2015起") & (fxr["固定效应"] == "无") & (fxr["变量"] == "FX_12M")].iloc[0]
fxall = fxr[(fxr["样本"] == "全样本") & (fxr["固定效应"] == "无") & (fxr["变量"] == "FX_12M")].iloc[0]
hl15 = hl[hl["样本"] == "2015起"].iloc[0]["半衰期(月)"]
lgP = lg[(lg["模型"] == "收窄12M(logit)") & (lg["变量"] == "P")].iloc[0]
lgz = lg[(lg["模型"] == "收窄12M(logit)") & (lg["变量"] == "z")].iloc[0]

# ---------------- 个股总表与结论 ----------------
M = dec.merge(sc[["名称", "个人/公募:汇率不变", "个人/公募:人民币升5%", "机构:汇率不变", "机构:人民币升5%",
                  "个人盈亏平衡汇率变化", "机构盈亏平衡汇率变化"]], on="名称") \
       .merge(cur[["名称", "收窄概率12M", "预期溢价变化12M", "历史月数"]], on="名称")


def call(base, stress):
    if stress > 0:
        return "买H（人民币升5%仍占优）"
    if base >= 0.01:
        return "偏H"
    if base > -0.01:
        return "无差别"
    return "买A"


M["个人/公募"] = [call(b, s) for b, s in zip(M["个人/公募:汇率不变"], M["个人/公募:人民币升5%"])]
M["保险/企业"] = [call(b, s) for b, s in zip(M["机构:汇率不变"], M["机构:人民币升5%"])]
M = M.sort_values("溢价_沪参考汇率", ascending=False)
short = set(M.loc[M["历史月数"] < 80, "名称"])

main_rows = []
for _, r in M.iterrows():
    nm = r["名称"] + ("<sup>†</sup>" if r["名称"] in short else "")
    main_rows.append([nm, r["行业"], pct(r["溢价_沪参考汇率"], 0), pct(r["溢价历史分位"], 0),
                      pp(r["个人:H税后−A(满1年)"]), pp(r["企业:H−A"]), pct(r["收窄概率12M"], 0),
                      pct(r["个人/公募:汇率不变"], 1, True), pct(r["个人/公募:人民币升5%"], 1, True),
                      ("升" + pct(-r["个人盈亏平衡汇率变化"], 1)) if r["个人盈亏平衡汇率变化"] < 0 else "需人民币贬值",
                      chip(r["个人/公募"]), chip(r["保险/企业"])])


def names(col, lab):
    return "、".join(M.loc[M[col].str.startswith(lab), "名称"])


# ---------------- 页面 ----------------
CSS = """
:root{
  /* 版式：单栏研究报告，左侧章节标记，宽表横向滚动 */
  --bg:#f6f7f9; --paper:#ffffff; --ink:#16202c; --muted:#5b6675; --rule:#d9dee5;
  --accent:#0d5c63; --accent-soft:#e3f0f0; --h:#1d3a8a; --h-soft:#e6ebf7; --a:#a3410b; --a-soft:#f7e9df;
  --pos:#0d6b3c; --neg:#b0302a; --plate:#ffffff;
  --f-display:"Noto Serif SC","Songti SC",serif; --f-body:"Noto Sans SC","PingFang SC","Microsoft YaHei",sans-serif;
  --f-num:"IBM Plex Mono",ui-monospace,monospace;
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#0f141a; --paper:#151c24; --ink:#e6ebf1; --muted:#9aa6b4; --rule:#2a3440;
  --accent:#5fb7bd; --accent-soft:#16312f; --h:#8fa9ee; --h-soft:#1b2440; --a:#e59866; --a-soft:#3a2416;
  --pos:#5fcf8f; --neg:#f08a80; --plate:#f4f5f7; color-scheme:dark}}
:root[data-theme="dark"]{
  --bg:#0f141a; --paper:#151c24; --ink:#e6ebf1; --muted:#9aa6b4; --rule:#2a3440;
  --accent:#5fb7bd; --accent-soft:#16312f; --h:#8fa9ee; --h-soft:#1b2440; --a:#e59866; --a-soft:#3a2416;
  --pos:#5fcf8f; --neg:#f08a80; --plate:#f4f5f7; color-scheme:dark}
body{background:var(--bg);color:var(--ink);font-family:var(--f-body);font-size:15px;line-height:1.75}
.wrap{max-width:980px;margin:0 auto;padding-block:40px 80px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
header.mast{border-bottom:2px solid var(--ink);padding-bottom:18px;margin-bottom:28px;display:grid;gap:6px}
.eyebrow{font-size:12px;letter-spacing:.12em;color:var(--accent);font-weight:600}
h1{font-family:var(--f-display);font-size:clamp(28px,5vw,42px);line-height:1.2;margin:0;text-wrap:balance}
.meta{color:var(--muted);font-size:13px;font-family:var(--f-num)}
h2{font-family:var(--f-display);font-size:24px;margin:56px 0 12px;padding-top:14px;border-top:1px solid var(--rule);text-wrap:balance;display:flex;gap:12px;align-items:baseline}
h2 .sec{font-family:var(--f-num);font-size:13px;color:var(--accent);font-weight:600}
h3{font-size:16px;margin:28px 0 8px}
p,li{max-width:68ch}
.lede{font-size:17px;max-width:70ch}
.key{background:var(--paper);border:1px solid var(--rule);border-left:4px solid var(--accent);padding:18px 22px;margin:22px 0;display:grid;gap:10px}
.key h3{margin:0;font-size:14px;letter-spacing:.08em;color:var(--accent)}
.key ol{margin:0;padding-left:1.3em;display:grid;gap:8px}
.key li{max-width:none}
.grid3{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin:18px 0}
.card{background:var(--paper);border:1px solid var(--rule);padding:14px 16px;display:grid;gap:6px;min-width:0}
.card .who{font-size:12px;letter-spacing:.08em;color:var(--muted)}
.card .what{font-family:var(--f-display);font-size:18px}
.card p{margin:0;font-size:13.5px}
.tw{overflow-x:auto;margin:14px 0;background:var(--paper);border:1px solid var(--rule)}
table{border-collapse:collapse;width:100%;font-size:13px;font-variant-numeric:tabular-nums}
th,td{padding:7px 10px;border-bottom:1px solid var(--rule);text-align:right;white-space:nowrap}
th:first-child,td:first-child{text-align:left}
th{font-weight:600;color:var(--muted);font-size:12px;background:var(--bg)}
td.pos{color:var(--pos)} td.neg{color:var(--neg)}
.tnote{font-size:12.5px;color:var(--muted);margin:-6px 0 14px;max-width:none}
figure{margin:18px 0;display:grid;gap:6px}
.plate{background:var(--plate);border:1px solid var(--rule);padding:8px}
.plate img{display:block;width:100%;height:auto}
figcaption{font-size:12.5px;color:var(--muted)}
.chip{display:inline-block;padding:1px 8px;border-radius:3px;font-size:12px;font-weight:600;white-space:nowrap}
.chip.h{background:var(--h-soft);color:var(--h)} .chip.hl{background:var(--h-soft);color:var(--h);opacity:.8}
.chip.n{background:var(--bg);color:var(--muted);border:1px solid var(--rule)} .chip.a{background:var(--a-soft);color:var(--a)}
.callout{border-left:3px solid var(--a);background:var(--a-soft);padding:10px 14px;margin:14px 0;font-size:14px}
.callout p{margin:0;max-width:none}
code{font-family:var(--f-num);font-size:12.5px}
sup{color:var(--a)}
footer{margin-top:60px;border-top:1px solid var(--rule);padding-top:14px;color:var(--muted);font-size:12.5px}
a{color:var(--accent)}
@media (max-width:600px){h2{font-size:20px} body{font-size:14.5px}}
"""

H = []
H.append(f"""<title>港股红利与AH抉择</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@600;700&family=Noto+Sans+SC:wght@400;600&family=IBM+Plex+Mono:wght@400;600&display=swap">
<style>{CSS}</style>
<div class="wrap">
<header class="mast">
  <div class="eyebrow">研究报告 · 港股高股息 / AH 溢价</div>
  <h1>港股红利与 AH 抉择</h1>
  <div class="meta">数据截至 2026-11-03 · 美联储 2026-09-16 重新加息 · 样本：恒生高股息率指数、23 只指标股中介持股、24 对高股息 AH 股</div>
</header>

<p class="lede">本报告回答三个问题：高利率环境下港股红利表现如何；谁在买港股高息股；对内地个人、公募和保险等机构，现在买 A 股还是买港股更划算。美联储加息和降息周期只作为辅助切分。</p>

<section class="key">
  <h3>核心结论</h3>
  <ol>
    <li><b>高利率并没有压制港股红利，反而相对占优。</b>美债 10Y ≥ 4% 的交易日，恒生高股息率全收益年化 {pct(A_hi['HSHDYI_年化'])}，恒指 {pct(A_hi['HSI_年化'])}，恒生科技 {pct(A_hi['HSTECH_年化'])}。在“利率高且维持不动”的阶段，高股息相对恒指的年化超额约 {pct(hrpA.loc['加息后维持','HSHDYI_年化超额'],0)}；一旦开始降息，资金转向高弹性资产，超额转为 {pct(hrpA.loc['降息调整段','HSHDYI_年化超额'],1,True)}。</li>
    <li><b>对美元资金的性价比已经很薄。</b>恒生高股息率股息率比美债 10Y 只高 {now['股息率-美债(bp)']:.0f}bp（历史 {pct(now['股息率-美债(bp)_分位'],0)} 分位），外资扣 10% 税后只剩 {now['外资税后-美债(bp)']:.0f}bp；但对内地资金，扣 20% 税后仍比中债 10Y 高 {now['内资税后-中债(bp)']:.0f}bp（{pct(now['内资税后-中债(bp)_分位'],0)} 分位）。</li>
    <li><b>边际买家是内地资金。</b>央企 H 股和红筹高息股的港股通持股占港股市值从 {gA_2017['港股通']:.1f}% 升到 {gA_now['港股通']:.1f}%，非港股通中介从 {gA_2017['非港股通']:.0f}% 降到 {gA_now['非港股通']:.0f}%；高利率月份换手更快。香港本地高息股被外资减持，但几乎没有内地资金接手。</li>
    <li><b>AH 溢价是“买哪边”的核心变量，门槛约 25–30%。</b>历史上溢价越高，未来 12 个月 H 股越跑赢 A 股，溢价也越可能收窄（半衰期约 {hl15:.0f} 个月）。溢价低于约 25% 时，即使股息免税的机构，买 H 也多数跑输。</li>
    <li><b>人民币升值是买 H 的主要风险。</b>过去 12 个月港币兑人民币 {pct(fx['最近12个月'],1,True)}。按汇率 1:1 传导做压力测试，人民币再升 5% 时，个人仍应买 H 的只剩 {names('个人/公募','买H')}；银行股在个人视角下基本没有优势。</li>
  </ol>
</section>

<div class="grid3">
  <div class="card"><div class="who">个人投资者 · 公募基金</div><div class="what">H 股股息 20% 税，门槛高</div>
    <p>买 H：{names('个人/公募','买H')}。偏 H：{names('个人/公募','偏H')}。其余买 A 或两边均可。</p></div>
  <div class="card"><div class="who">保险公司 · 其他居民企业（持有满 12 个月）</div><div class="what">两边股息都免税，看溢价回归</div>
    <p>买 H：{names('保险/企业','买H')}。偏 H：{names('保险/企业','偏H')}。买 A：{names('保险/企业','买A')}。</p></div>
  <div class="card"><div class="who">指数层面</div><div class="what">A 股红利略好（个人）</div>
    <p>AH 溢价指数 {ii['AH指数']:.1f}（{pct(ii['AH分位'],0)} 分位）。港股通高股息股息率 {ii['HK高股息DY']:.2f}% × 0.8 = {ii['HK高股息DY']*0.8:.2f}%，低于中证红利 {ii['A红利DY']:.2f}%。</p></div>
</div>
""")

# 1 高利率
ep_rows = [[f"{r['起']} 至 {r['止']}", f"{r['美债10Y均值']:.2f}%", pct(r["HSHDYI"], 1, True), pct(r["HSI"], 1, True),
            pct(r["HSHDYI_超额"], 1, True), pct(r["HSTECH"], 1, True) if pd.notna(r["HSTECH"]) else "—"] for _, r in ep.iterrows()]
ph_rows = []
for ph in ["加息后维持", "降息调整段", "降息后维持"]:
    if ph in hrpA.index:
        r = hrpA.loc[ph]
        ph_rows.append([ph, int(r["交易日"]), pct(r["HSHDYI_年化"], 1, True), pct(r["HSI_年化"], 1, True), pct(r["HSHDYI_年化超额"], 1, True)])
evh = ev[ev["指数"] == "HSHDYI"].dropna(subset=["T+250"])
ev_rows = [[f"{r['T0']} {r['方向']}", r["标签"], pct(r["T+60"], 1, True), pct(r["T+120"], 1, True), pct(r["T+250"], 1, True)] for _, r in evh.iterrows()]
H.append(f"""
<h2><span class="sec">01</span>高利率环境下的港股红利</h2>
<p>以美债 10Y ≥ 4% 定义高利率（另用联邦基金利率 ≥ 4%、实际利率 ≥ 1.5% 交叉验证，方向一致）。2007 年以来只有 3 段较长的高利率时期：</p>
{table(["高利率区间", "美债10Y均值", "恒生高股息率", "恒指", "高股息超额", "恒生科技"], ep_rows, note="全收益口径，港币。恒生高股息率指数 2012-12 前为回溯数据。")}
<p>高利率期间按美联储阶段细分（美债 ≥ 4%，年化）：</p>
{table(["联储阶段", "交易日", "恒生高股息率", "恒指", "高股息年化超额"], ph_rows)}
<p>首次加息或降息后 HSHDYI 相对恒指的超额（交易日）：加息后红利占优，降息后风险偏好回升、红利相对落后。</p>
{table(["首次调整", "类型", "T+60", "T+120", "T+250"], ev_rows)}
{img("fig1_相对强弱与周期.png", "图1 恒生高股息率相对恒指（全收益）与美联储周期")}
""")

# 2 利差
sp_rows = []
for k in ["C05加息 2015-12", "C08加息 2022-03", "C09降息 2024-09", "C10加息 2026-09", "当前"]:
    r = snap.loc[k]
    sp_rows.append([k, f"{r['股息率']:.2f}%", f"{r['美债10Y']:.2f}%", f"{r['中债10Y']:.2f}%", f"{r['股息率-美债(bp)']:.0f}bp",
                    f"{r['外资税后-美债(bp)']:.0f}bp", f"{r['内资税后-中债(bp)']:.0f}bp"])
q_rows = [[r["利差五分位"], f"{r['利差均值']:.0f}bp", pct(r["未来12M超额均值"], 1, True), pct(r["胜率"], 0)] for _, r in q5.iterrows()]
H.append(f"""
<h2><span class="sec">02</span>利率绝对水平与股息率利差</h2>
<p>这一轮加息和 2015、2022 年两轮最大的不同在于起点：前两轮开始时股息率比美债高 400bp 以上，这一轮只高几十个基点。</p>
{table(["时点", "HSHDYI股息率", "美债10Y", "中债10Y", "股息率−美债", "外资税后−美债", "内资税后−中债"], sp_rows,
       note="外资按 H 股预提税 10%、港股通个人按 20% 近似。周期时点为当月月末值。")}
<p>按“股息率 − 美债 10Y”分五组，起点利差越高，未来 12 个月相对恒指的超额越高（月度重叠样本，描述性）：</p>
{table(["利差五分位", "平均利差", "未来12个月超额", "胜率"], q_rows)}
<p>月度回归中，联邦基金利率变动本身几乎没有解释力；中债 10Y 每下行 100bp，高股息相对恒指当月超额约 +5%（t≈−3.3），比美债更重要。当前美债利差已低于历史上任何有后续数据的月份，用历史关系推断属于外推。</p>
{img("fig3_股息率利差.png", "图2 恒生高股息率股息率相对美债、中债的利差")}
""")

# 3 资金
H.append(f"""
<h2><span class="sec">03</span>谁在买：中介持股</h2>
<p>23 只指标股，Wind 四类中介口径（港股通 / 国际 / 本地 / 中资），2015-09 至 2026-10 月末。主指标用“非港股通中介合计”，以抵消中介之间的托管转移。</p>
{table(["", "高利率月均（美债≥4%）", "非高利率月均"], [
    ["A 组 港股通净买入（亿港元）", num(ccA.loc['高利率(美债≥4%)','月均港股通净买入亿'],0), num(ccA.loc['非高利率','月均港股通净买入亿'],0)],
    ["A 组 非港股通中介净买入", num(ccA.loc['高利率(美债≥4%)','月均非港股通净买入亿'],0), num(ccA.loc['非高利率','月均非港股通净买入亿'],0)],
    ["B 组 港股通净买入", num(ccB.loc['高利率(美债≥4%)','月均港股通净买入亿'],1), num(ccB.loc['非高利率','月均港股通净买入亿'],1)],
    ["B 组 非港股通中介净买入", num(ccB.loc['高利率(美债≥4%)','月均非港股通净买入亿'],0), num(ccB.loc['非高利率','月均非港股通净买入亿'],0)],
], note="A 组：央企 H 股/红筹高息股 12 只；B 组：香港本地高息股 8 只。净买入 = 持股变化 × 月均价，估算值。")}
<ul>
<li>面板回归（个股固定效应、按月聚类）：高利率月份非港股通中介持股每月多降 {pan_hi['系数']:.2f}pp（t={pan_hi['t值']:.1f}），港股通每月多增 {pan_sb['系数']:.2f}pp（t={pan_sb['t值']:.1f}）。</li>
<li>港股通只对“股息率 − 中债”利差敏感；非港股通中介对“股息率 − 美债”利差为正、对中债利差为负。</li>
<li>2024-10 至 2026-08，B 组非港股通中介净卖出约 {abs(eB['非港股通累计净买入亿']):,.0f} 亿港元，港股通只买入 {eB['港股通累计净买入亿']:,.0f} 亿。</li>
</ul>
{img("fig6_分组中介持股.png", "图3 指标股中介持股结构（红色阴影为美债10Y≥4%）")}
""")

# 4 AH
b_rows = [[r["溢价档"], int(r["样本"]), pct(r["平均溢价"], 0), pct(r["H减A_企业"], 1, True), pct(r["H减A_个人"], 1, True),
           pct(r["个人H胜率"], 0), pct(r["溢价12M变化"], 1, True)] for _, r in b41.iterrows()]
i_rows = [[r["AH档"], int(r["月数"]), pct(r["HK减A_企业"], 1, True), pct(r["HK减A_个人"], 1, True), pct(r["个人HK胜率"], 0)] for _, r in i46.iterrows()]
H.append(f"""
<h2><span class="sec">04</span>AH 溢价：历史规律</h2>
<p>溢价 = A 价 ÷（H 价 × 汇率）− 1。同一公司 A、H 每股派息相同，所以 H 股息率 = A 股息率 × (1 + 溢价)。收益为人民币全收益，H 股含汇率变化。</p>
{table(["起点溢价", "样本", "平均溢价", "未来12M H−A（税前）", "H−A（个人税后）", "个人H胜率", "溢价12M变化"], b_rows,
       note=f"24 对，月度。不重叠年度样本（每年 12 月起点，{int(rb['年数'])} 年）：溢价每高 10pp，H−A 多约 {rb['系数_每1倍溢价']*10:.1f}pp，t={rb['t值']:.1f}。")}
<p>指数层面（恒生 AH 溢价指数分档 → 未来 12 个月 中证港股通高股息 − 中证红利，2014-11 起）：</p>
{table(["AH 指数", "月数", "差值（税前）", "差值（个人税后）", "个人港股胜率"], i_rows)}
{img("fig9_AH溢价历史.png", "图4 恒生AH溢价指数与24对高息AH股溢价中位数")}
{img("fig11_溢价分档_未来12M.png", "图5 起点溢价与未来12个月 H−A")}
""")

# 5 税
H.append("""
<h2><span class="sec">05</span>股息税：不同投资者的门槛</h2>
""" + table(["投资者", "A 股股息", "港股通 H 股股息", "只看股息的盈亏平衡溢价"], [
    ["个人，A 股持有满 1 年", "0%", "20%", "25%"],
    ["个人，A 股持有 1 个月 ~ 1 年", "10%", "20%", "12.5%"],
    ["公募基金（含港股通基金）", "按持有期：满 1 年 0%", "20%（比照个人）", "25% / 12.5%"],
    ["保险公司、居民企业，直接持有满 12 个月", "免税", "免税（H 股公司不代扣）", "0%"],
    ["居民企业，不满 12 个月", "计入企业所得税", "自行申报企业所得税", "约 0%（两边对称）"],
    ["红筹股（移动、中海油），企业经港股通", "—", "H 股免税条款不适用，扣税可否抵免需逐只核验", "按 10% 预扣情景，H 仍占优"],
], note="依据：财税〔2014〕81 号、〔2016〕127 号，企业所得税法实施条例第 83 条。险资通过资管产品、委托或 QDII 间接持有不适用直接持有的免税条款；QDII 持 H 股一般按 10% 代扣。") + """
<div class="callout"><p>对免税机构，股息上的门槛是 0，但总收益上的门槛仍在 25–30%：溢价低时往往继续扩大，价格损失超过 H 股多出的股息。回归中 H−A（税前）为零的溢价约 24%（2015 起月度）到 31%（不重叠年度）。</p></div>
""")

# 6 汇率
c53_rows = []
for (_, r), (_, n) in zip(c53.iterrows(), c53n.iterrows()):
    c53_rows.append([r["汇率档"]] + [f"{pct(r[c],1,True)}（{int(n[c])}）" for c in ["<25%", "25-40%", ">40%"]])
H.append(f"""
<h2><span class="sec">06</span>人民币升值情景</h2>
<p>港币挂钩美元，人民币升值会直接压低 H 股的人民币收益。过去 12 个月港币兑人民币 {pct(fx['最近12个月'],1,True)}。2007 年以来 12 个月变化的分布：中位数 {pct(fx['0.5'],1,True)}，25% 分位 {pct(fx['0.25'],1,True)}，10% 分位 {pct(fx['0.1'],1,True)}，5% 分位 {pct(fx['0.05'],1,True)}；人民币 12 个月升值超过 3% 的时间占 {pct(fx['人民币升值>3%的频率'],0)}，超过 5% 占 {pct(fx['人民币升值>5%的频率'],0)}。</p>
<h3>历史上汇率传导并不完全</h3>
<p>回归“未来 12 个月 H−A ~ 起点溢价 + 汇率变化”，汇率系数 2015 年起为 {fx15['系数']:.2f}（t={fx15['t值']:.1f}），全样本 {fxall['系数']:.2f}（t={fxall['t值']:.1f}），都不显著。人民币走强时，内地资产往往同时上涨，H 股的港币价格部分抵消了汇率损失。但分组看，高溢价股票在人民币大幅升值时的超额确实明显缩小：</p>
{table(["港币兑人民币 12 个月", "溢价<25%", "溢价25–40%", "溢价>40%"], c53_rows, note="单元格为未来 12 个月 H−A（税前）均值，括号为样本数，2015 年起。")}
<h3>压力测试：汇率 1:1 传导</h3>
<p>为保守起见，情景按“H 股港币价格不因汇率变化”计算：预期 H−A = 溢价模型预测 + 汇率变化，个人和公募再扣 20% × H 股息率。下表“可承受升值”是预期 H−A 降为 0 时人民币还能升值多少。</p>
{img("fig12_汇率情景.png", "图6 各汇率情景下个人/公募的预期 H−A")}
""")

# 7 收窄
n_rows = [[r["溢价档"], int(r["样本"]), pct(r["收窄概率6M"], 0), pct(r["收窄概率12M"], 0), pct(r["收窄超10pp概率"], 0),
           pct(r["中位变化12M"], 1, True)] for _, r in nar.iterrows()]
H.append(f"""
<h2><span class="sec">07</span>溢价收窄概率</h2>
{table(["起点溢价", "样本", "6个月内收窄", "12个月内收窄", "12个月收窄超10pp", "12个月溢价变化中位数"], n_rows,
       note="24 对，月度重叠样本，频率为历史经验值。")}
<ul>
<li>溢价水平本身是主要预测变量：logit 中溢价系数 t={lgP['t值']:.1f}；相对本股历史均值的偏离（z）加入后不显著（t={lgz['t值']:.1f}）。</li>
<li>面板 AR(1) 估计的均值回归半衰期约 {hl15:.0f} 个月（2015 年起），即溢价偏离本股均值的部分大约一年回归一半。</li>
<li>溢价低于 25% 时，12 个月内收窄的概率不到一半，且中位数是继续扩大，这是低溢价时买 H 跑输的原因。</li>
</ul>
{img("fig13_溢价收窄概率.png", "图7 溢价收窄的历史频率与当前个股概率")}
""")

# 8 个股总表
H.append(f"""
<h2><span class="sec">08</span>当前个股：买 A 还是买 H</h2>
<p>判定规则（预期 H−A 为溢价模型 + 汇率 1:1 传导）：人民币升 5% 时预期 H−A 仍为正 → 买H；汇率不变时 ≥ +1% → 偏H；−1% 到 +1% → 无差别；≤ −1% → 买A。</p>
{table(["股票", "行业", "溢价", "本股历史分位", "个人税后股息差", "机构股息差", "12M收窄概率", "个人预期H−A·汇率不变", "个人·人民币升5%", "个人可承受升值", "个人/公募", "保险/企业"],
       main_rows, cls="wide",
       note="溢价按沪港通参考汇率（买 H 实际汇率）。股息率为 Wind 一致预期 DPS。† A 股上市晚于 2021 年，历史不足 7 年，分位与模型参考性较弱。机构列的预期 H−A 未扣税。")}
{img("fig10_当前AH个股.png", "图8 当前 AH 溢价与个人税后股息率差")}
""")

# 9 局限
H.append("""
<h2><span class="sec">09</span>局限</h2>
<ul>
<li>高利率样本集中在 2023–2026 年，与内地资金大规模南下重叠，难以完全区分利率与资金面的影响。</li>
<li>中介不等于最终投资者：中资和本地中介都混有内外资客户，托管转移会被误认为买卖；Wind 在 2020-02 至 2021-01 有大规模的中介归类调整。</li>
<li>一致预期 DPS 大多从 2015 年起，税后回测只覆盖 2015 年之后；红筹股港股通税率需逐只核验。</li>
<li>收益回归和收窄概率使用月度重叠样本，t 值偏高；不重叠年度样本只有约 10 年。两地没有套利机制，高溢价可以长期维持。</li>
<li>汇率情景为压力测试口径，历史上实际传导明显小于 1:1。</li>
</ul>
<footer>
<p>数据：Wind、Bloomberg、FRED、港交所、恒生指数公司、中证指数公司（经用户整理）。分析脚本：<code>analysis/</code> 下 fed_cycle_hk_dividend.py、spread_and_flows.py、high_rate_regime.py、ccass_holdings.py、ah_premium.py、ah_scenarios.py；本页由 build_report.py 从 output/ 结果表生成。本报告为研究分析，不构成投资建议。</p>
</footer>
</div>
""")

DST.write_text("\n".join(H), encoding="utf-8")
print("written", DST, round(DST.stat().st_size / 1e6, 2), "MB")
