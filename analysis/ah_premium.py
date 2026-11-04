"""AH 溢价与“买 A 还是买港股高股息”。

数据: data/AH溢价与高股息研究数据_20261104.xlsx（24 对 AH 股，月末面板），
      data/hk_dividend_fed_cycle_20261103.xlsx（港股红利指数、美债/中债、汇率）。

口径
  溢价 P = A价 / (H价 × 汇率) − 1。历史用人民币中间价（全样本可得）；当前决策用沪港通参考汇率（买H用银行卖出价）。
  同一公司 A/H 每股派息相同（H股按汇率折港币派发），故 H 股息率(人民币) = A 股息率 × (1+P)。
  股息率用一致预期 DPS（人民币，Wind FY1）/ 价格。
  税：个人 A 股持有满1年 0%、1个月~1年 10%；港股通个人 H 股 20%；
      企业/险资持有满12个月 A、H 均按免税处理（H 股以财税〔2014〕81号为依据，需按主体核验）。
  收益：A、H 后复权价（含税前股息再投资），H 再乘汇率变化折人民币；个人 H 股扣 20% 股息税近似为 0.2×H股息率。

运行: python3 analysis/ah_premium.py
"""
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
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
F_AH = ROOT / "data" / "AH溢价与高股息研究数据_20261104.xlsx"
F_MAIN = ROOT / "data" / "hk_dividend_fed_cycle_20261103.xlsx"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)

TAX_H_IND = 0.20          # 港股通个人 H 股
TAX_A_IND_SHORT = 0.10    # A 股个人持有 1个月~1年
SECTOR = {"建设银行": "银行", "工商银行": "银行", "中国银行": "银行", "农业银行": "银行", "交通银行": "银行",
          "招商银行": "银行", "中信银行": "银行", "民生银行": "银行", "光大银行": "银行",
          "中国石油": "能源", "中国石化": "能源", "中国海洋石油": "能源", "中国神华": "能源", "兖矿能源": "能源", "中煤能源": "能源",
          "中国移动": "电信", "中国电信": "电信", "华能国际": "公用/基建", "江苏宁沪高速": "公用/基建",
          "中国平安": "保险", "中国人寿": "保险", "中国太保": "保险", "中远海控": "周期", "海螺水泥": "周期"}

xl = pd.read_excel(F_AH, sheet_name=["AH月末", "配对与股本", "指数日度", "预期DPS月度", "中国利率"], header=4)
P = xl["AH月末"].copy()
P["日期"] = pd.to_datetime(P["日期"])
P["月"] = P["日期"].dt.to_period("M")
P = P[P["采样期"] == "月末最后共同交易日"].copy()  # 当月未结束的单独做当前快照
cur = xl["AH月末"].copy()
cur["日期"] = pd.to_datetime(cur["日期"])
cur = cur[cur["日期"] == cur["日期"].max()].copy()

pairs = xl["配对与股本"]
dps = xl["预期DPS月度"].copy()
dps["观测日"] = pd.to_datetime(dps["观测日"])
dps = dps[dps["Wind代码"].str.endswith((".SH", ".SZ"))]
dps["月"] = dps["观测日"].dt.to_period("M")
dps = dps.groupby(["Wind代码", "月"])["一致预期DPS(CNY)"].last().rename("DPS").reset_index()


def attach(df):
    df = df.merge(dps, left_on=["A代码", "月"], right_on=["Wind代码", "月"], how="left").drop(columns="Wind代码")
    df["DPS"] = df.groupby("A代码")["DPS"].ffill(limit=2)
    df["A股息率"] = df["DPS"] / df["A收盘(CNY)"]
    df["H股息率"] = df["DPS"] / (df["H收盘(HKD)"] * df["中间价(CNY/HKD)"])  # 人民币口径
    df["行业"] = df["名称"].map(SECTOR)
    return df


P = attach(P.sort_values(["A代码", "日期"]))
cur["月"] = cur["日期"].dt.to_period("M")
cur = attach(cur)

# ---------- 前瞻 12 个月收益（人民币） ----------
g = P.groupby("A代码")
P["A_12M"] = g["A后复权价"].shift(-12) / P["A后复权价"] - 1
P["H_12M_HKD"] = g["H后复权价"].shift(-12) / P["H后复权价"] - 1
P["FX_12M"] = g["中间价(CNY/HKD)"].shift(-12) / P["中间价(CNY/HKD)"] - 1
P["H_12M"] = (1 + P["H_12M_HKD"]) * (1 + P["FX_12M"]) - 1
P["H_12M_个人税后"] = P["H_12M"] - TAX_H_IND * P["H股息率"]
P["H减A_税前"] = P["H_12M"] - P["A_12M"]          # 企业/险资视角
P["H减A_个人"] = P["H_12M_个人税后"] - P["A_12M"]  # 个人（A持有满1年）
P["溢价变化12M"] = g["溢价(中间价)"].shift(-12) - P["溢价(中间价)"]
# 个股自身历史分位（只用当时及以前的数据，避免前视）
P["溢价分位_至当时"] = g["溢价(中间价)"].transform(lambda s: s.expanding(24).apply(lambda x: (x <= x[-1]).mean(), raw=True))
P.to_csv(OUT / "40_AH月度面板.csv", index=False, encoding="utf-8-sig")

# ---------- 1. 市场层面：AH 溢价指数历史 ----------
idx = xl["指数日度"].copy()
idx["日期"] = pd.to_datetime(idx["日期"])
idx = idx.set_index("日期").sort_index()
ahi = idx["AH溢价指数"].dropna()
pool_med = P.groupby("日期")["溢价(中间价)"].median()

# ---------- 2. 溢价分档 -> 未来 12 个月 H−A（人民币） ----------
B = P.dropna(subset=["H减A_个人", "溢价(中间价)"]).copy()
B["溢价档"] = pd.cut(B["溢价(中间价)"], [-1, 0, 0.15, 0.25, 0.40, 0.60, 5],
                    labels=["<0", "0-15%", "15-25%", "25-40%", "40-60%", ">60%"])
t_b = B.groupby("溢价档", observed=True).agg(
    样本=("H减A_个人", "size"), 平均溢价=("溢价(中间价)", "mean"),
    H减A_企业=("H减A_税前", "mean"), H减A_个人=("H减A_个人", "mean"),
    个人H胜率=("H减A_个人", lambda s: (s > 0).mean()), 其中汇率=("FX_12M", "mean"),
    溢价12M变化=("溢价变化12M", "mean"))
t_b.to_csv(OUT / "41_溢价分档_未来12M_H减A.csv", encoding="utf-8-sig")

B["分位档"] = pd.cut(B["溢价分位_至当时"], [0, .2, .4, .6, .8, 1.0001], labels=["0-20%", "20-40%", "40-60%", "60-80%", "80-100%"])
t_q = B.dropna(subset=["分位档"]).groupby("分位档", observed=True).agg(
    样本=("H减A_个人", "size"), H减A_企业=("H减A_税前", "mean"), H减A_个人=("H减A_个人", "mean"),
    个人H胜率=("H减A_个人", lambda s: (s > 0).mean()), 溢价12M变化=("溢价变化12M", "mean"))
t_q.to_csv(OUT / "42_个股自身溢价分位_未来12M_H减A.csv", encoding="utf-8-sig")

# 面板回归（个股固定效应，按月聚类；重叠样本）
B["t"] = B["日期"].rank(method="dense").astype(int)
rows = []
for y in ["H减A_税前", "H减A_个人", "溢价变化12M"]:
    for fml in [f"{y} ~ Q('溢价(中间价)')", f"{y} ~ Q('溢价(中间价)') + C(A代码)"]:
        dd = B.dropna(subset=[y])
        r = smf.ols(fml, dd).fit(cov_type="cluster", cov_kwds={"groups": dd["t"]})
        rows.append(dict(因变量=y, 固定效应="个股" if "C(" in fml else "无", 系数_每1倍溢价=r.params["Q('溢价(中间价)')"],
                         t值=r.tvalues["Q('溢价(中间价)')"], N=int(r.nobs), R2=r.rsquared))
pd.DataFrame(rows).to_csv(OUT / "43_溢价对未来12M的回归.csv", index=False, encoding="utf-8-sig")

# ---------- 3. 当前个股决策表 ----------
hist = P.groupby("A代码")["溢价(中间价)"]
c = cur.set_index("A代码")
dec = pd.DataFrame({
    "名称": c["名称"], "行业": c["行业"],
    "溢价_沪参考汇率": c["溢价(沪参考)"], "溢价_中间价": c["溢价(中间价)"],
    "溢价历史分位": [(hist.get_group(k) <= v).mean() if k in hist.groups else np.nan for k, v in c["溢价(中间价)"].items()],
    "溢价历史中位": [hist.get_group(k).median() if k in hist.groups else np.nan for k in c.index],
    "历史起点": [P.loc[P["A代码"] == k, "日期"].min().date() for k in c.index],
    "A股息率": c["A股息率"],
})
# 当前决策按沪参考汇率（实际买 H 成本）
dec["H股息率(人民币)"] = c["DPS"] / (c["H收盘(HKD)"] * c["沪参考买H汇率"])
dec["H个人税后"] = dec["H股息率(人民币)"] * (1 - TAX_H_IND)
dec["个人:H税后−A(满1年)"] = dec["H个人税后"] - dec["A股息率"]
dec["个人:H税后−A(1月-1年)"] = dec["H个人税后"] - dec["A股息率"] * (1 - TAX_A_IND_SHORT)
dec["企业:H−A"] = dec["H股息率(人民币)"] - dec["A股息率"]


def verdict(r):
    if pd.isna(r["个人:H税后−A(满1年)"]):
        return "缺DPS"
    d = r["个人:H税后−A(满1年)"]
    if d > 0.005:
        return "个人也偏H"
    if d > 0:
        return "个人略偏H（优势<0.5pp，汇率可抵消）"
    return "个人偏A；企业/险资偏H" if r["溢价_沪参考汇率"] > 0 else "两类都偏A"


dec["结论(股息维度)"] = dec.apply(verdict, axis=1)
dec = dec.sort_values("溢价_沪参考汇率", ascending=False)
dec.to_csv(OUT / "44_当前AH个股决策表.csv", encoding="utf-8-sig")

# ---------- 4. 指数层面：中证红利 vs 中证港股通高股息（人民币全收益） ----------
mx = pd.read_excel(F_MAIN, sheet_name=["指数日度", "估值与利差", "解释变量"])
hk = mx["指数日度"].set_index("日期")["中证港股通高股息人民币全收益(CNY)"]
hk_dy = mx["估值与利差"].set_index("日期")["CSI930914近12个月股息率"]
ex = mx["解释变量"].set_index("日期").sort_index()
I = pd.DataFrame({"A红利TR": idx["中证红利全收益"], "A红利低波TR": idx["红利低波全收益"], "A红利DY": idx["中证红利股息率(%)"],
                  "AH指数": idx["AH溢价指数"]}).join(pd.DataFrame({"HK高股息TR": hk, "HK高股息DY": hk_dy}), how="outer").sort_index()
I = I.resample("ME").last()
I["US10Y"] = ex["美国10年期国债恒定期限收益率"].resample("ME").last()
I["CN10Y"] = ex["中国10年期国债Wind重取对照"].resample("ME").last()
I["HK12M"] = I["HK高股息TR"].shift(-12) / I["HK高股息TR"] - 1
I["A12M"] = I["A红利TR"].shift(-12) / I["A红利TR"] - 1
I["HK减A_企业"] = I["HK12M"] - I["A12M"]
I["HK减A_个人"] = I["HK12M"] - TAX_H_IND * I["HK高股息DY"] / 100 - I["A12M"]
I["DY差_个人"] = I["HK高股息DY"] * (1 - TAX_H_IND) - I["A红利DY"]
I["AH分位"] = I["AH指数"].expanding(24).apply(lambda x: (x <= x[-1]).mean(), raw=True)
I.to_csv(OUT / "45_指数层面_A红利vs港股通高股息.csv", encoding="utf-8-sig")
J = I.dropna(subset=["HK减A_个人", "AH指数"]).copy()
J["AH档"] = pd.cut(J["AH指数"], [0, 115, 125, 135, 145, 300], labels=["<115", "115-125", "125-135", "135-145", "≥145"])
J["利率"] = np.where(J["US10Y"] >= 4, "美债≥4%", "美债<4%")
t_i = J.groupby("AH档", observed=True).agg(月数=("HK减A_个人", "size"), HK减A_企业=("HK减A_企业", "mean"),
                                          HK减A_个人=("HK减A_个人", "mean"), 个人HK胜率=("HK减A_个人", lambda s: (s > 0).mean()))
t_i.to_csv(OUT / "46_AH指数分档_未来12M_港股高股息减A红利.csv", encoding="utf-8-sig")
t_r = J.groupby(["利率", "AH档"], observed=True).agg(月数=("HK减A_个人", "size"), HK减A_个人=("HK减A_个人", "mean"))
t_r.to_csv(OUT / "47_利率x AH指数_未来12M.csv".replace(" ", ""), encoding="utf-8-sig")

# ---------- 图 ----------
fig, ax = plt.subplots(2, 1, figsize=(13, 8), sharex=True, gridspec_kw={"height_ratios": [1.6, 1]})
ax[0].plot(ahi.index, ahi, color="#1f3b57", lw=1.1, label="恒生AH溢价指数")
for q, ls in [(0.2, ":"), (0.5, "--"), (0.8, ":")]:
    ax[0].axhline(ahi.quantile(q), color="gray", ls=ls, lw=.8)
    ax[0].text(ahi.index[0], ahi.quantile(q), f" {int(q*100)}%分位 {ahi.quantile(q):.0f}", color="gray", fontsize=8, va="bottom")
ax[0].plot(pool_med.index, 100 + pool_med * 100, color="#e39b2d", lw=1.1, label="24对高息AH 溢价中位数(折算为100+溢价%)")
hi = I["US10Y"] >= 4
for a_ in ax:
    a_.fill_between(I.index, 0, 1, where=hi.values, transform=a_.get_xaxis_transform(), color="#d9534f", alpha=.12, lw=0)
ax[0].legend(loc="upper left")
ax[0].set_title("AH 溢价历史（红色阴影=美债10Y≥4%）")
ax[1].plot(I.index, I["DY差_个人"], color="#08306b", label="港股通高股息股息率×0.8 − 中证红利股息率 (pp)")
ax[1].axhline(0, color="black", lw=.6)
ax[1].legend(loc="upper left")
ax[1].set_xlim(pd.Timestamp("2007-01-01"), I.index.max())
fig.tight_layout()
fig.savefig(OUT / "fig9_AH溢价历史.png", dpi=140)

fig, ax = plt.subplots(figsize=(11, 7))
d = dec.dropna(subset=["个人:H税后−A(满1年)"]).sort_values("溢价_沪参考汇率")
colors = ["#08306b" if v > 0 else "#e39b2d" for v in d["个人:H税后−A(满1年)"]]
ax.barh(d["名称"], d["溢价_沪参考汇率"] * 100, color=colors)
ax.axvline(25, color="#d9534f", ls="--")
ax.text(25.5, -0.6, "个人盈亏平衡 25%", color="#d9534f", fontsize=9)
for i_, (p, dd, q) in enumerate(zip(d["溢价_沪参考汇率"], d["个人:H税后−A(满1年)"], d["溢价历史分位"])):
    ax.text(p * 100 + 1, i_, f"税后股息差 {dd*100:+.2f}pp | 历史分位 {q:.0%}", va="center", fontsize=8)
ax.set_xlabel("AH 溢价（%，沪港通参考汇率）")
ax.set_title(f"当前 AH 溢价与个人投资者税后股息率差（{cur['日期'].max().date()}；深蓝=个人买H股息更高）")
fig.tight_layout()
fig.savefig(OUT / "fig10_当前AH个股.png", dpi=140)

fig, axes = plt.subplots(1, 2, figsize=(13, 4.6))
t_b[["H减A_企业", "H减A_个人"]].mul(100).plot.bar(ax=axes[0], color=["#9ecae1", "#08306b"])
axes[0].set_title("个股：起点溢价 → 未来12个月 H−A（人民币，%）")
axes[0].axhline(0, color="black", lw=.6)
t_i[["HK减A_企业", "HK减A_个人"]].mul(100).plot.bar(ax=axes[1], color=["#fdd0a2", "#e39b2d"])
axes[1].set_title("指数：AH指数分档 → 未来12个月 港股通高股息−中证红利（%）")
axes[1].axhline(0, color="black", lw=.6)
for a_ in axes:
    a_.tick_params(axis="x", rotation=0)
fig.tight_layout()
fig.savefig(OUT / "fig11_溢价分档_未来12M.png", dpi=140)
print("done", cur["日期"].max().date())

# ---------- 稳健性：不重叠样本（每年12月起点），按年聚类 ----------
NB = B[B["日期"].dt.month == 12].copy()
rows = []
for y in ["H减A_税前", "H减A_个人", "溢价变化12M"]:
    dd = NB.dropna(subset=[y])
    for fe in (False, True):
        fml = f"{y} ~ Q('溢价(中间价)')" + (" + C(A代码)" if fe else "")
        r = smf.ols(fml, dd).fit(cov_type="cluster", cov_kwds={"groups": dd["日期"].dt.year})
        rows.append(dict(因变量=y, 固定效应="个股" if fe else "无", 系数_每1倍溢价=r.params["Q('溢价(中间价)')"],
                         t值=r.tvalues["Q('溢价(中间价)')"], N=int(r.nobs), 年数=dd["日期"].dt.year.nunique()))
pd.DataFrame(rows).to_csv(OUT / "48_稳健性_不重叠年度样本.csv", index=False, encoding="utf-8-sig")
