"""利率绝对水平 / 股息率利差 / 资金面 对港股红利的影响。

1. 美债10Y绝对水平分档 × 美联储阶段 -> HSHDYI 未来6/12个月绝对收益与超额
2. 股息率-美债10Y、股息率-中债10Y 利差分档 -> 未来收益；HAC(12) 预测回归
3. 不同投资者视角的税后利差（外资 vs 港股通内资）
4. 南向资金与利差的关系：边际买家是谁

运行: python3 analysis/spread_and_flows.py [xlsx路径]
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

for _f in ["/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"]:
    if Path(_f).exists():
        font_manager.fontManager.addfont(_f)
plt.rcParams["font.sans-serif"] = ["WenQuanYi Zen Hei", "Noto Sans CJK SC", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

ROOT = Path(__file__).resolve().parents[1]
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "hk_dividend_fed_cycle_20261103.xlsx"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)

xl = pd.read_excel(XLSX, sheet_name=["月度全收益", "解释变量", "估值与利差", "联储日度利率"])
m = xl["月度全收益"].set_index("月份")
m = m[m["月份状态"] == "完整月"]
ex = xl["解释变量"].set_index("日期").sort_index()
val = xl["估值与利差"].set_index("日期").sort_index()
fed = xl["联储日度利率"].set_index("日期").sort_index()

# ---------- 月度面板（月末值） ----------
P = pd.DataFrame(index=m.index)
P["HSHDYI"] = m["恒生高股息率全收益月末值"]
P["HSI"] = m["恒生指数全收益月末值"]
P["CSI"] = m["中证港股通高股息港元全收益月末值"]
em = ex.resample("MS").last()  # 月初标签 = 当月最后观测
P["US10Y"] = em["美国10年期国债恒定期限收益率"]
P["US10Y_real"] = em["美国10年期实际国债恒定期限收益率"]
P["CN10Y"] = em["中国10年期国债Wind重取对照"]
P["HIBOR1M"] = em["HIBOR一个月"]
P["USDCNH"] = em["美元兑离岸人民币Bloomberg长历史"]
P["南向净买入"] = ex["南向合计每日净买入"].resample("MS").sum(min_count=1)  # 亿港元
vm = val.resample("MS").last()
P["DY"] = vm["HSHDYI近12个月股息率"]
P["PE"] = vm["HSHDYI滚动市盈率"]
P["利差_US"] = P["DY"] - P["US10Y"]           # pp
P["利差_CN"] = P["DY"] - P["CN10Y"]
P["利差_US实际"] = P["DY"] - P["US10Y_real"]
# 税后视角（近似）：H股/红筹股息预提税——外资(非居民企业/个人)约10%，港股通内地个人H股20%。
# 成份股以H股和红筹为主，这里统一用近似税率，仅作量级参考。
P["外资税后利差"] = P["DY"] * 0.9 - P["US10Y"]
P["外资税后对HIBOR"] = P["DY"] * 0.9 - P["HIBOR1M"]
P["内资税后利差"] = P["DY"] * 0.8 - P["CN10Y"]
phase = fed["阶段（按FRED记录日）"].resample("MS").last()
P["阶段"] = phase.reindex(P.index).str.replace("（含段内暂停）", "", regex=False)

for h in (6, 12):
    P[f"未来{h}M_绝对"] = P["HSHDYI"].shift(-h) / P["HSHDYI"] - 1
    P[f"未来{h}M_超额"] = P[f"未来{h}M_绝对"] - (P["HSI"].shift(-h) / P["HSI"] - 1)
P.to_csv(OUT / "10_月度面板_利差与资金.csv", encoding="utf-8-sig")

# ---------- 1. 美债10Y绝对水平 × 阶段（2007-07 起，HSHDYI 含 2012 前回溯） ----------
D = P.loc["2007-07-01":].copy()
D["美债10Y档"] = pd.cut(D["US10Y"], [0, 2, 3, 4, 10], labels=["<2%", "2-3%", "3-4%", "≥4%"], right=False)
agg = dict(月数=("未来12M_超额", "count"), 未来12M绝对=("未来12M_绝对", "mean"),
           未来12M超额=("未来12M_超额", "mean"), 超额胜率=("未来12M_超额", lambda s: (s > 0).mean()))
t1 = D.groupby("美债10Y档", observed=True).agg(**agg)
t1b = D.groupby(["美债10Y档", "阶段"], observed=True).agg(**agg)
t1.to_csv(OUT / "11_美债10Y水平分档.csv", encoding="utf-8-sig")
t1b.to_csv(OUT / "11b_美债10Y水平x联储阶段.csv", encoding="utf-8-sig")

# ---------- 2. 利差分档（股息率数据 2012-12 起） ----------
S = P.dropna(subset=["DY"]).copy()
S["利差US档"] = pd.cut(S["利差_US"], [-10, 1.5, 3, 4.5, 20], labels=["<150bp", "150-300bp", "300-450bp", "≥450bp"], right=False)
S["利差CN档"] = pd.cut(S["利差_CN"], [-10, 2.5, 3.5, 4.5, 20], labels=["<250bp", "250-350bp", "350-450bp", "≥450bp"], right=False)
t2 = S.groupby("利差US档", observed=True).agg(**agg, 平均PE=("PE", "mean"))
t3 = S.groupby("利差CN档", observed=True).agg(**agg, 平均PE=("PE", "mean"))
t23 = S.groupby(["利差US档", "利差CN档"], observed=True).agg(**agg)
t2.to_csv(OUT / "12_美债利差分档.csv", encoding="utf-8-sig")
t3.to_csv(OUT / "13_中债利差分档.csv", encoding="utf-8-sig")
t23.to_csv(OUT / "14_美债x中债利差交叉.csv", encoding="utf-8-sig")

# 预测回归：未来12M超额/绝对 ~ 利差（重叠样本，HAC 12 阶）
rows = []
for y in ("未来12M_超额", "未来12M_绝对"):
    for xs in (["利差_US"], ["利差_CN"], ["利差_US", "利差_CN"], ["外资税后利差", "内资税后利差"], ["US10Y", "CN10Y", "DY"]):
        d = S[[y] + xs].dropna()
        r = sm.OLS(d[y], sm.add_constant(d[xs])).fit(cov_type="HAC", cov_kwds={"maxlags": 12})
        for v in xs:
            rows.append(dict(因变量=y, 模型="+".join(xs), 变量=v, 系数=r.params[v], t值=r.tvalues[v], N=int(r.nobs), R2=r.rsquared))
pd.DataFrame(rows).to_csv(OUT / "15_利差预测回归.csv", index=False, encoding="utf-8-sig")

# ---------- 3. 各周期起点/当前：各视角利差 ----------
pts = {"C05加息 2015-12": "2015-12-01", "C06降息 2019-08": "2019-08-01", "C07降息 2020-03": "2020-03-01",
       "C08加息 2022-03": "2022-03-01", "C09降息 2024-09": "2024-09-01", "C10加息 2026-09": "2026-09-01"}
vd = val.ffill()
cur_date = vd.index.max()
snap = []
for k, d in list(pts.items()) + [("当前", str(cur_date.date()))]:
    r = P.loc[:d].iloc[-1] if k != "当前" else None
    if k == "当前":
        dy = vd["HSHDYI近12个月股息率"].iloc[-1]
        e = ex.ffill().iloc[-1]
        us, cn, hb, real = e["美国10年期国债恒定期限收益率"], e["中国10年期国债Wind重取对照"], e["HIBOR一个月"], e["美国10年期实际国债恒定期限收益率"]
    else:
        dy, us, cn, hb, real = r["DY"], r["US10Y"], r["CN10Y"], r["HIBOR1M"], r["US10Y_real"]
    snap.append({"时点": k, "股息率": dy, "美债10Y": us, "美债实际10Y": real, "HIBOR1M": hb, "中债10Y": cn,
                 "股息率-美债(bp)": (dy - us) * 100, "外资税后-美债(bp)": (dy * .9 - us) * 100,
                 "外资税后-HIBOR(bp)": (dy * .9 - hb) * 100, "内资税后-中债(bp)": (dy * .8 - cn) * 100})
snap = pd.DataFrame(snap).set_index("时点")
# 历史分位
for c in ["股息率-美债(bp)", "外资税后-美债(bp)", "内资税后-中债(bp)"]:
    base = {"股息率-美债(bp)": S["利差_US"], "外资税后-美债(bp)": S["外资税后利差"], "内资税后-中债(bp)": S["内资税后利差"]}[c] * 100
    snap[c + "_分位"] = snap[c].apply(lambda v: (base.dropna() <= v).mean())
snap.to_csv(OUT / "16_各视角利差快照.csv", encoding="utf-8-sig")

# ---------- 4. 南向资金：边际买家对利差的反应 ----------
F = P.loc["2015-01-01":].dropna(subset=["南向净买入", "利差_US", "利差_CN"]).copy()
F["ΔUSDCNH%"] = P["USDCNH"].pct_change().reindex(F.index) * 100
rows = []
for xs in (["利差_CN"], ["利差_US"], ["利差_CN", "利差_US", "ΔUSDCNH%"]):
    d = F[["南向净买入"] + xs].dropna()
    r = sm.OLS(d["南向净买入"], sm.add_constant(d[xs])).fit(cov_type="HAC", cov_kwds={"maxlags": 6})
    for v in xs:
        rows.append(dict(模型="+".join(xs), 变量=v, 系数_亿港元每pp=r.params[v], t值=r.tvalues[v], N=int(r.nobs), R2=r.rsquared))
pd.DataFrame(rows).to_csv(OUT / "17_南向资金对利差回归.csv", index=False, encoding="utf-8-sig")
F.groupby("阶段")["南向净买入"].agg(["mean", "median", "count"]).to_csv(OUT / "18_各阶段月均南向净买入.csv", encoding="utf-8-sig")

# ---------- 图 ----------
fig, ax = plt.subplots(1, 2, figsize=(13, 5))
for a, xcol, lab in [(ax[0], "利差_US", "股息率 − 美债10Y (pp)"), (ax[1], "利差_CN", "股息率 − 中债10Y (pp)")]:
    d = S[[xcol, "未来12M_超额"]].dropna()
    a.scatter(d[xcol], d["未来12M_超额"] * 100, s=14, alpha=.6, color="#1f3b57")
    b = np.polyfit(d[xcol], d["未来12M_超额"] * 100, 1)
    xx = np.linspace(d[xcol].min(), d[xcol].max(), 50)
    a.plot(xx, np.polyval(b, xx), color="#d9534f")
    cur = S[xcol].dropna().iloc[-1] if xcol == "利差_CN" else snap.loc["当前", "股息率-美债(bp)"] / 100
    a.axvline(cur, color="#e39b2d", ls="--")
    a.text(cur, a.get_ylim()[1] * .9, " 当前", color="#e39b2d")
    a.axhline(0, color="black", lw=.6)
    a.set_xlabel(lab)
    a.set_ylabel("HSHDYI 未来12个月超额恒指 (%)")
fig.suptitle("起点利差与未来12个月超额（月度，2012-12 起，重叠样本）")
fig.tight_layout()
fig.savefig(OUT / "fig4_利差与未来超额.png", dpi=140)

fig, ax1 = plt.subplots(figsize=(13, 4.5))
ax1.bar(F.index, F["南向净买入"], width=25, color="#9ecae1", label="南向月度净买入(亿港元)")
ax2 = ax1.twinx()
ax2.plot(F.index, F["利差_CN"], color="#e39b2d", label="股息率−中债10Y")
ax2.plot(F.index, F["利差_US"], color="#1f3b57", label="股息率−美债10Y")
ax1.set_ylabel("亿港元")
ax2.set_ylabel("pp")
h1, l1 = ax1.get_legend_handles_labels()
h2, l2 = ax2.get_legend_handles_labels()
ax1.legend(h1 + h2, l1 + l2, loc="upper left")
ax1.set_title("南向资金与股息率利差")
fig.tight_layout()
fig.savefig(OUT / "fig5_南向资金与利差.png", dpi=140)
print("done")
