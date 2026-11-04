"""AH：人民币升值情景 + 溢价收窄概率。

依赖 analysis/ah_premium.py 生成的 output/40_AH月度面板.csv 与 44_当前AH个股决策表.csv。

1. 汇率情景
   历史回归：未来12个月 H−A（人民币，税前） = a + b·起点溢价 + c·港币兑人民币12个月变化 + 个股固定效应
   c 衡量汇率传导（c=1 表示人民币升值1%，H 相对 A 少赚1%；c<1 表示港币计价的 H 股价格部分抵消）。
   情景：港币兑人民币 12 个月变化 +3% / 0 / −3% / −5% / −8%（负数=人民币升值）。
   个人（公募）再扣 20%×H股息率；企业/险资不扣。
   另给出每只股票“盈亏平衡升值幅度”：预期 H−A 恰为 0 时人民币需升值多少。

2. 溢价收窄概率
   事件：未来 6 / 12 个月溢价下降（收窄），以及下降超过 10 个百分点。
   (a) 经验频率：按起点溢价档 × 个股自身历史分位档；
   (b) Logit：收窄 ~ 起点溢价 + 相对自身历史均值的偏离（z）+ 当期汇率变化方向不用（避免前视），按月聚类；
   (c) 均值回归半衰期：面板 AR(1)（个股固定效应）。
   当前个股给出 12 个月收窄概率、预期溢价变化。

运行: python3 analysis/ah_scenarios.py
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
OUT = ROOT / "output"
TAX_H = 0.20
FX_SCEN = {"港币升3%(人民币贬)": 0.03, "汇率不变": 0.0, "人民币升3%": -0.03, "人民币升5%": -0.05, "人民币升8%": -0.08}

P = pd.read_csv(OUT / "40_AH月度面板.csv", parse_dates=["日期"])
DEC = pd.read_csv(OUT / "44_当前AH个股决策表.csv")
P = P.sort_values(["A代码", "日期"])
P["P"] = P["溢价(中间价)"]
g = P.groupby("A代码")
P["P_6M后"] = g["P"].shift(-6)
P["P_12M后"] = g["P"].shift(-12)
P["Δ6"] = P["P_6M后"] - P["P"]
P["Δ12"] = P["P_12M后"] - P["P"]
# 自身历史偏离（只用当时及以前，至少 24 个月）
P["均值_至当时"] = g["P"].transform(lambda s: s.expanding(24).mean())
P["标准差_至当时"] = g["P"].transform(lambda s: s.expanding(24).std())
P["z"] = (P["P"] - P["均值_至当时"]) / P["标准差_至当时"]
P["t"] = P["日期"].rank(method="dense").astype(int)

# ---------- 0. 汇率历史分布 ----------
fx = P.groupby("日期")["中间价(CNY/HKD)"].first().sort_index()
fx12 = fx.pct_change(12).dropna()
fx_dist = pd.DataFrame({"分位": [.05, .1, .25, .5, .75, .9, .95]})
fx_dist["港币兑人民币12M变化"] = fx_dist["分位"].map(fx12.quantile)
fx_dist.loc[len(fx_dist)] = ["最近12个月", fx12.iloc[-1]]
fx_dist.loc[len(fx_dist)] = ["人民币升值>3%的频率", (fx12 < -0.03).mean()]
fx_dist.loc[len(fx_dist)] = ["人民币升值>5%的频率", (fx12 < -0.05).mean()]
fx_dist.to_csv(OUT / "50_港币兑人民币12M变化分布.csv", index=False, encoding="utf-8-sig")

# ---------- 1. 汇率情景回归 ----------
R = P.dropna(subset=["H减A_税前", "P", "FX_12M"]).copy()
R15 = R[R["日期"] >= "2015-01-01"]
reg_rows = []
models = {}
for nm, d in [("全样本", R), ("2015起", R15)]:
    for fe in (False, True):
        fml = "H减A_税前 ~ P + FX_12M" + (" + C(A代码)" if fe else "")
        r = smf.ols(fml, d).fit(cov_type="cluster", cov_kwds={"groups": d["t"]})
        models[(nm, fe)] = r
        for v in ["Intercept", "P", "FX_12M"]:
            if v in r.params:
                reg_rows.append(dict(样本=nm, 固定效应="个股" if fe else "无", 变量=v, 系数=r.params[v], t值=r.tvalues[v],
                                     N=int(r.nobs), R2=r.rsquared))
# 汇率对溢价本身的影响
for nm, d in [("全样本", R), ("2015起", R15)]:
    dd = d.dropna(subset=["Δ12"])
    r = smf.ols("Δ12 ~ P + FX_12M + C(A代码)", dd).fit(cov_type="cluster", cov_kwds={"groups": dd["t"]})
    for v in ["P", "FX_12M"]:
        reg_rows.append(dict(样本=nm, 固定效应="个股", 变量=f"[溢价变化12M] {v}", 系数=r.params[v], t值=r.tvalues[v],
                             N=int(r.nobs), R2=r.rsquared))
pd.DataFrame(reg_rows).to_csv(OUT / "51_汇率情景回归.csv", index=False, encoding="utf-8-sig")

# 用 2015 起、无固定效应的模型做情景（个股固定效应会把各股历史平均溢价回归的特殊性外推，A 股新上市股票样本短）
m = models[("2015起", False)]
a, b, c = m.params["Intercept"], m.params["P"], m.params["FX_12M"]
cur = DEC.copy()
cur["P"] = cur["溢价_中间价"]
rows = []
for _, r_ in cur.iterrows():
    rec = {"名称": r_["名称"], "行业": r_["行业"], "当前溢价": r_["溢价_沪参考汇率"], "H股息率(人民币)": r_["H股息率(人民币)"]}
    tax = TAX_H * r_["H股息率(人民币)"]
    for k, v in FX_SCEN.items():
        # 机械传导（压力测试）：H 股港币价格不因汇率变化，汇率 1:1 进入人民币收益
        mech = a + b * r_["P"] + 1.0 * v
        emp_ = a + b * r_["P"] + c * v   # 历史经验传导
        rec[f"机构:{k}"] = mech
        rec[f"个人/公募:{k}"] = mech - tax
        rec[f"经验·个人/公募:{k}"] = emp_ - tax
    # 机械传导下的盈亏平衡汇率变化：a + bP + fx (− tax) = 0
    rec["机构盈亏平衡汇率变化"] = -(a + b * r_["P"])
    rec["个人盈亏平衡汇率变化"] = -(a + b * r_["P"] - tax)
    rows.append(rec)
S = pd.DataFrame(rows).sort_values("当前溢价", ascending=False)
S.to_csv(OUT / "52_汇率情景_个股预期H减A.csv", index=False, encoding="utf-8-sig")

# 实际历史：按汇率方向分组的 H−A
R15b = R15.copy()
R15b["汇率档"] = pd.cut(R15b["FX_12M"], [-1, -0.05, -0.02, 0.02, 0.05, 1],
                      labels=["人民币升>5%", "人民币升2-5%", "±2%", "港币升2-5%", "港币升>5%"])
R15b["溢价档"] = pd.cut(R15b["P"], [-1, 0.25, 0.4, 5], labels=["<25%", "25-40%", ">40%"])
fx_hist = R15b.pivot_table(index="汇率档", columns="溢价档", values="H减A_税前", aggfunc="mean", observed=True)
fx_cnt = R15b.pivot_table(index="汇率档", columns="溢价档", values="H减A_税前", aggfunc="size", observed=True)
fx_hist.to_csv(OUT / "53_历史_汇率x溢价_H减A均值.csv", encoding="utf-8-sig")
fx_cnt.to_csv(OUT / "53b_历史_汇率x溢价_样本数.csv", encoding="utf-8-sig")

# ---------- 2. 溢价收窄概率 ----------
N = P.dropna(subset=["Δ12", "P"]).copy()
N["收窄12M"] = (N["Δ12"] < 0).astype(int)
N["收窄10pp"] = (N["Δ12"] < -0.10).astype(int)
N["溢价档"] = pd.cut(N["P"], [-1, 0, 0.15, 0.25, 0.4, 0.6, 5], labels=["<0", "0-15%", "15-25%", "25-40%", "40-60%", ">60%"])
N["z档"] = pd.cut(N["z"], [-10, -1, 0, 1, 10], labels=["低于均值1σ以上", "均值下方", "均值上方", "高于均值1σ以上"])
emp = N.groupby("溢价档", observed=True).agg(样本=("收窄12M", "size"), 收窄概率12M=("收窄12M", "mean"),
                                         收窄超10pp概率=("收窄10pp", "mean"), 平均变化12M=("Δ12", "mean"),
                                         中位变化12M=("Δ12", "median"))
emp6 = P.dropna(subset=["Δ6"]).assign(档=lambda x: pd.cut(x["P"], [-1, 0, 0.15, 0.25, 0.4, 0.6, 5],
                                                         labels=["<0", "0-15%", "15-25%", "25-40%", "40-60%", ">60%"]))
emp["收窄概率6M"] = emp6.groupby("档", observed=True)["Δ6"].apply(lambda s: (s < 0).mean())
emp.to_csv(OUT / "54_溢价档_收窄概率.csv", encoding="utf-8-sig")
empz = N.dropna(subset=["z档"]).groupby("z档", observed=True).agg(样本=("收窄12M", "size"), 收窄概率12M=("收窄12M", "mean"),
                                                                 收窄超10pp概率=("收窄10pp", "mean"), 平均变化12M=("Δ12", "mean"))
empz.to_csv(OUT / "55_自身偏离_收窄概率.csv", encoding="utf-8-sig")
cross = N.dropna(subset=["z档"]).pivot_table(index="溢价档", columns="z档", values="收窄12M", aggfunc="mean", observed=True)
cross.to_csv(OUT / "56_溢价档x自身偏离_收窄概率.csv", encoding="utf-8-sig")

L = N.dropna(subset=["z"]).copy()
logit = smf.logit("收窄12M ~ P + z", L).fit(disp=0, cov_type="cluster", cov_kwds={"groups": L["t"]})
logit10 = smf.logit("收窄10pp ~ P + z", L).fit(disp=0, cov_type="cluster", cov_kwds={"groups": L["t"]})
ols_d = smf.ols("Δ12 ~ P + z", L).fit(cov_type="cluster", cov_kwds={"groups": L["t"]})
pd.DataFrame([dict(模型=n, 变量=v, 系数=r.params[v], t值=r.tvalues[v], N=int(r.nobs))
              for n, r in [("收窄12M(logit)", logit), ("收窄超10pp(logit)", logit10), ("溢价变化12M(OLS)", ols_d)]
              for v in r.params.index]).to_csv(OUT / "57_收窄概率模型.csv", index=False, encoding="utf-8-sig")

# 半衰期：ΔP_t = α_i + β P_{t-1}
P["P_lag"] = g["P"].shift()
P["dP"] = P["P"] - P["P_lag"]
ar = smf.ols("dP ~ P_lag + C(A代码)", P.dropna(subset=["dP", "P_lag"])).fit()
beta = ar.params["P_lag"]
half_life = np.log(0.5) / np.log(1 + beta) if -1 < beta < 0 else np.nan
ar15 = smf.ols("dP ~ P_lag + C(A代码)", P[P["日期"] >= "2015-01-01"].dropna(subset=["dP", "P_lag"])).fit()
beta15 = ar15.params["P_lag"]
hl15 = np.log(0.5) / np.log(1 + beta15) if -1 < beta15 < 0 else np.nan
pd.DataFrame([{"样本": "全样本", "β": beta, "半衰期(月)": half_life},
              {"样本": "2015起", "β": beta15, "半衰期(月)": hl15}]).to_csv(OUT / "58_溢价均值回归半衰期.csv", index=False, encoding="utf-8-sig")

# 当前个股：用最新月末的 z（与历史同口径），溢价用当前中间价口径
last = P.sort_values("日期").groupby("A代码").tail(1).set_index("A代码")
mu = P.groupby("A代码")["P"].mean()
sd = P.groupby("A代码")["P"].std()
cur = DEC.set_index("A代码").copy()
cur["P"] = cur["溢价_中间价"]
cur["z"] = (cur["P"] - mu.reindex(cur.index)) / sd.reindex(cur.index)
cur["收窄概率12M"] = logit.predict(cur[["P", "z"]])
cur["收窄超10pp概率"] = logit10.predict(cur[["P", "z"]])
cur["预期溢价变化12M"] = ols_d.predict(cur[["P", "z"]])
cur["经验频率(同溢价档)"] = pd.cut(cur["P"], [-1, 0, 0.15, 0.25, 0.4, 0.6, 5],
                              labels=["<0", "0-15%", "15-25%", "25-40%", "40-60%", ">60%"]).map(emp["收窄概率12M"]).astype(float)
cur["历史月数"] = P.groupby("A代码").size().reindex(cur.index)
out = cur[["名称", "行业", "溢价_沪参考汇率", "溢价历史分位", "z", "收窄概率12M", "收窄超10pp概率", "预期溢价变化12M",
           "经验频率(同溢价档)", "历史月数"]].sort_values("收窄概率12M", ascending=False)
out.to_csv(OUT / "59_当前个股_溢价收窄概率.csv", encoding="utf-8-sig")

# ---------- 图 ----------
scen_cols = list(FX_SCEN)
fig, ax = plt.subplots(figsize=(12, 8))
S2 = S.sort_values("当前溢价")
yy = np.arange(len(S2))
colors = ["#7f2704", "#bdbdbd", "#6baed6", "#2171b5", "#08306b"]
for k, col in zip(scen_cols, colors):
    ax.scatter(S2[f"个人/公募:{k}"] * 100, yy, color=col, s=26, label=k, zorder=3)
ax.axvline(0, color="black", lw=.8)
ax.set_yticks(yy)
ax.set_yticklabels([f"{n}（溢价{p*100:.0f}%）" for n, p in zip(S2["名称"], S2["当前溢价"])], fontsize=9)
ax.set_xlabel("模型预期 未来12个月 H−A（人民币，个人/公募税后，%）")
ax.set_title("人民币汇率情景下的 H−A 预期（溢价回归 + 汇率 1:1 机械传导，压力测试口径）")
ax.legend(loc="lower right", fontsize=9)
ax.grid(axis="x", alpha=.3)
fig.tight_layout()
fig.savefig(OUT / "fig12_汇率情景.png", dpi=140)

fig, axes = plt.subplots(1, 2, figsize=(13, 4.8))
emp[["收窄概率6M", "收窄概率12M", "收窄超10pp概率"]].mul(100).plot.bar(ax=axes[0], color=["#c6dbef", "#4292c6", "#08306b"])
axes[0].set_title("起点溢价档 → 之后溢价收窄的历史频率（%）")
axes[0].axhline(50, color="gray", ls="--", lw=.8)
axes[0].tick_params(axis="x", rotation=0)
o = out.sort_values("收窄概率12M")
axes[1].barh(o["名称"], o["收窄概率12M"] * 100, color=["#08306b" if v >= .6 else "#9ecae1" for v in o["收窄概率12M"]])
axes[1].axvline(50, color="gray", ls="--", lw=.8)
axes[1].set_title("当前个股：未来12个月溢价收窄概率（logit：溢价水平+自身偏离）")
axes[1].tick_params(axis="y", labelsize=8)
fig.tight_layout()
fig.savefig(OUT / "fig13_溢价收窄概率.png", dpi=140)
print("done; FX coef c =", round(c, 3), "half-life(2015+) =", round(hl15, 1))
