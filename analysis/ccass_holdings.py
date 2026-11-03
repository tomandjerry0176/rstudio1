"""高利率环境下港股高息股的中介持股变化（23只指标股，Wind 四类中介口径）。

分组: A=央企H股/红筹高息股, B=香港本地/传统外资高息股, C=非红利对照
持有人: 港股通 / 国际中介 / 本地中介 / 中资中介
  * Wind 在 2020-02 ~ 2021-01 期间国际与本地中介之间存在大规模归类迁移（几乎所有股票同步反向跳变），
    因此主指标用"国际+本地"（非内资中介），国际单列时剔除该窗口。
  * 领展(0823) 2018-10~2019-01、2023-02 原始比例异常，置空；领展未纳入港股通。
输出: output/3x_*.csv, fig6~fig8

运行: python3 analysis/ccass_holdings.py
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
F_HOLD = ROOT / "data" / "港股23只股票_四类中介持股与配套数据_20261103.xlsx"
F_MAIN = ROOT / "data" / "hk_dividend_fed_cycle_20261103.xlsx"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)

GROUP_NAME = {"A": "A 央企H股/红筹高息", "B": "B 香港本地高息", "C": "C 非红利对照"}
CATS = {"港股通": "港股通合计持股", "国际": "国际中介持股", "本地": "本地中介持股", "中资": "中资中介持股"}
RECLASS = (pd.Timestamp("2020-02-01"), pd.Timestamp("2021-01-31"))

# ---------- 读取 ----------
h = pd.read_excel(F_HOLD, sheet_name="四类持股", header=4)
mm = pd.read_excel(F_HOLD, sheet_name="月度配套", header=4)
h["日期"] = pd.to_datetime(h["日期"])
mm["日期"] = pd.to_datetime(mm["日期"])

# 领展异常值置空
bad = (h["Wind代码"] == "0823.HK") & (h["日期"].between("2018-10-01", "2019-01-31") | h["日期"].between("2023-02-01", "2023-02-28"))
h.loc[bad, list(CATS.values())] = np.nan
h["非内资"] = h["国际中介持股"] + h["本地中介持股"]
# 非港股通中介合计：国际+本地+中资，可抵消中介之间的托管转移（如 1288 2026-09 国际→中资）
h["非港股通"] = h["非内资"] + h["中资中介持股"]

ex = pd.read_excel(F_MAIN, sheet_name="解释变量").set_index("日期").sort_index()
fed = pd.read_excel(F_MAIN, sheet_name="联储日度利率").set_index("日期").sort_index()


def rate_on(dates, s):
    return s.dropna().reindex(pd.DatetimeIndex(dates), method="ffill").values


# ---------- 月末面板 ----------
M = h[h["采样标签"].str.contains("month_end")].copy()
M = M.merge(mm[["日期", "Wind代码", "收盘价(HKD)", "滚动股息率Wind(%)", "预期股息率(%)", "自由流通股数Wind"]],
            on=["日期", "Wind代码"], how="left")
M = M.sort_values(["Wind代码", "日期"])
M["US10Y"] = rate_on(M["日期"], ex["美国10年期国债恒定期限收益率"])
M["CN10Y"] = rate_on(M["日期"], ex["中国10年期国债Wind重取对照"])
M["FFR"] = rate_on(M["日期"], fed["统一上界(%)"])
M["阶段"] = pd.Series(fed["阶段（按FRED记录日）"].reindex(pd.DatetimeIndex(M["日期"]), method="ffill").values,
                    index=M.index).str.replace("（含段内暂停）", "", regex=False)
M["高利率"] = np.where(M["US10Y"] >= 4, "高利率(美债≥4%)", "非高利率")
M["股息率"] = M["预期股息率(%)"].fillna(M["滚动股息率Wind(%)"])  # 优先一致预期
M["利差_US"] = M["股息率"] * 0.9 - M["US10Y"]   # 外资近似税后
M["利差_CN"] = M["股息率"] * 0.8 - M["CN10Y"]   # 港股通个人H股近似税后
# 分母用港股上市股数（H股=H股股数）：Wind 自由流通股数存在跳变，不宜作分母
M["流通股"] = M["香港上市股数/单位数"]

g = M.groupby("Wind代码")
for k, col in list(CATS.items()) + [("非内资", "非内资"), ("非港股通", "非港股通")]:
    M[f"{k}_占流通%"] = M[col] / M["流通股"] * 100
    # 月度净买入估算(亿港元) = Δ持股 × 月均价
    dsh = g[col].diff()
    px = (M["收盘价(HKD)"] + g["收盘价(HKD)"].shift()) / 2
    M[f"{k}_净买入亿"] = dsh * px / 1e8
    M[f"Δ{k}_pp"] = g[f"{k}_占流通%"].diff()
# 国际单列：剔除归类迁移窗口（含跨窗口的那一次差分）
win = M["日期"].between(*RECLASS) | (M["日期"] == pd.Timestamp("2021-01-29"))
for c in ["国际_净买入亿", "Δ国际_pp", "本地_净买入亿", "Δ本地_pp"]:
    M.loc[win, c] = np.nan
# 港股通从无数据到有数据的第一期不是买入
first_sb = M["港股通合计持股"].notna() & g["港股通合计持股"].shift().isna()
M.loc[first_sb, ["港股通_净买入亿", "Δ港股通_pp"]] = np.nan
M.to_csv(OUT / "30_中介持股月度面板.csv", index=False, encoding="utf-8-sig")

# ---------- 1. 分组：持股占自由流通市值比例（市值加权） ----------
M["流通市值"] = M["流通股"] * M["收盘价(HKD)"]  # 港股上市市值
rows = []
for (d, grp), x in M.groupby(["日期", "研究分组"]):
    rec = dict(日期=d, 分组=grp)
    for k, col in list(CATS.items()) + [("非内资", "非内资"), ("非港股通", "非港股通")]:
        v = x[col] * x["收盘价(HKD)"]
        ok = v.notna()
        rec[k] = v[ok].sum() / x.loc[ok, "流通市值"].sum() * 100 if ok.any() else np.nan
    rows.append(rec)
G = pd.DataFrame(rows)
G.loc[G["日期"].between(*RECLASS), ["国际", "本地"]] = np.nan
G.to_csv(OUT / "31_分组持股比例_港股市值加权.csv", index=False, encoding="utf-8-sig")

# ---------- 2. 高利率 vs 非高利率：月均净买入与占比变化 ----------
agg_cols = {f"{k}_净买入亿": "sum" for k in ["港股通", "非内资", "国际", "中资", "非港股通"]}
F = M.groupby(["日期", "研究分组"]).agg({**agg_cols, **{f"Δ{k}_pp": "mean" for k in ["港股通", "非港股通", "非内资", "国际", "中资"]},
                                       "高利率": "first", "阶段": "first", "US10Y": "first"}).reset_index()
F = F[F["日期"] >= "2015-10-01"]
t2 = F.groupby(["研究分组", "高利率"]).agg(月数=("US10Y", "size"),
                                         **{f"月均{k}净买入亿": (f"{k}_净买入亿", "mean") for k in ["港股通", "非港股通", "非内资", "国际", "中资"]},
                                         **{f"月均Δ{k}pp": (f"Δ{k}_pp", "mean") for k in ["港股通", "非港股通", "非内资", "中资"]})
t2.to_csv(OUT / "32_高利率vs非高利率_中介净买入.csv", encoding="utf-8-sig")
t2b = F[F["高利率"].str.startswith("高")].groupby(["研究分组", "阶段"]).agg(
    月数=("US10Y", "size"), **{f"月均{k}净买入亿": (f"{k}_净买入亿", "mean") for k in ["港股通", "非港股通", "非内资", "中资"]})
t2b.to_csv(OUT / "33_高利率内按联储阶段.csv", encoding="utf-8-sig")

# 高利率区间累计
EPIS = {"2023-07~2024-07 高利率(加息后维持)": ("2023-07-01", "2024-07-31"),
        "2024-10~2026-08 高利率(降息中)": ("2024-10-01", "2026-08-31"),
        "2026-09~ 重新加息": ("2026-09-01", "2026-11-30"),
        "对照 2019-08~2022-02 低利率": ("2019-08-01", "2022-02-28")}
rows = []
for name, (a, b) in EPIS.items():
    x = F[F["日期"].between(a, b)]
    for grp, y in x.groupby("研究分组"):
        g0 = G[(G["分组"] == grp) & (G["日期"] < a)].tail(1)
        g1 = G[(G["分组"] == grp) & (G["日期"] <= b)].tail(1)
        rows.append(dict(区间=name, 分组=grp, 月数=len(y),
                         港股通累计净买入亿=y["港股通_净买入亿"].sum(), 非内资累计净买入亿=y["非内资_净买入亿"].sum(),
                         中资累计净买入亿=y["中资_净买入亿"].sum(), 非港股通累计净买入亿=y["非港股通_净买入亿"].sum(),
                         港股通占比变化pp=(g1["港股通"].values - g0["港股通"].values)[0] if len(g0) and len(g1) else np.nan,
                         非内资占比变化pp=(g1["非内资"].values - g0["非内资"].values)[0] if len(g0) and len(g1) else np.nan))
pd.DataFrame(rows).to_csv(OUT / "34_高利率区间_累计中介净买入.csv", index=False, encoding="utf-8-sig")

# ---------- 3. 面板回归：持股变化 ~ 个股税后利差 ----------
R = M[(M["日期"] >= "2017-06-01")].copy()
R["L利差_US"] = R.groupby("Wind代码")["利差_US"].shift()
R["L利差_CN"] = R.groupby("Wind代码")["利差_CN"].shift()
R["ΔUS10Y"] = R.groupby("Wind代码")["US10Y"].diff()
R["高"] = (R["US10Y"] >= 4).astype(int)
R["t"] = R["日期"].rank(method="dense").astype(int)
rows = []
for y in ["Δ非港股通_pp", "Δ非内资_pp", "Δ港股通_pp"]:
    for fml in [f"Q('{y}') ~ L利差_US + C(Wind代码)", f"Q('{y}') ~ L利差_CN + C(Wind代码)",
                f"Q('{y}') ~ L利差_US + L利差_CN + C(Wind代码)", f"Q('{y}') ~ 高 + C(Wind代码)"]:
        for sub, cond in [("A+B 高息股", R["研究分组"].isin(["A", "B"])), ("A 组", R["研究分组"] == "A"), ("B 组", R["研究分组"] == "B")]:
            d = R[cond].dropna(subset=[y] + [v for v in ["L利差_US", "L利差_CN"] if v in fml])
            r = smf.ols(fml, d).fit(cov_type="cluster", cov_kwds={"groups": d["t"]})
            for v in ["L利差_US", "L利差_CN", "高"]:
                if v in r.params:
                    rows.append(dict(因变量=y, 样本=sub, 模型=fml.split("~")[1].replace("+ C(Wind代码)", "+个股固定效应"),
                                     变量=v, 系数=r.params[v], t值=r.tvalues[v], N=int(r.nobs)))
pd.DataFrame(rows).to_csv(OUT / "35_面板回归_持股变化对利差.csv", index=False, encoding="utf-8-sig")

# ---------- 4. 首次加/降息前后（周度采样）：事件前最后一个采样 -> +3个月 ----------
EV = {"2015-12 加息": "2015-12-16", "2019-07 降息": "2019-07-31", "2020-03 降息": "2020-03-03",
      "2022-03 加息": "2022-03-16", "2024-09 降息": "2024-09-18", "2025-12 末次降息": "2025-12-10", "2026-09 再加息": "2026-09-16"}
H = h.copy()
H["流通股"] = H["香港上市股数/单位数"]
rows = []
for name, dd in EV.items():
    dd = pd.Timestamp(dd)
    for code, x in H.groupby("Wind代码"):
        x = x.sort_values("日期")
        pre = x[x["日期"] <= dd].tail(1)
        post = x[(x["日期"] > dd) & (x["日期"] <= dd + pd.Timedelta(days=95))].tail(1)
        if pre.empty or post.empty:
            continue
        rec = dict(事件=name, 代码=code, 名称=x["名称"].iloc[0], 分组=x["研究分组"].iloc[0], 起=pre["日期"].iloc[0].date(), 止=post["日期"].iloc[0].date())
        for k, col in [("港股通", "港股通合计持股"), ("非港股通", "非港股通"), ("非内资", "非内资"), ("中资", "中资中介持股")]:
            a_, b_ = pre[col].iloc[0], post[col].iloc[0]
            rec[f"Δ{k}_占流通pp"] = (b_ - a_) / post["流通股"].iloc[0] * 100 if pd.notna(a_) and pd.notna(b_) else np.nan
        # 跨归类迁移窗口时非内资仍可用（国际+本地合计），无需剔除
        rows.append(rec)
E = pd.DataFrame(rows)
E.to_csv(OUT / "36_首次调整前后3个月_个股持股变化.csv", index=False, encoding="utf-8-sig")
E.groupby(["事件", "分组"])[[c for c in E.columns if c.startswith("Δ")]].mean().to_csv(
    OUT / "37_首次调整前后3个月_分组均值.csv", encoding="utf-8-sig")

# ---------- 5. 当前快照 ----------
last = M["日期"].max()
S = M[M["日期"] == last].set_index("Wind代码")
P12 = M[M["日期"] == M[M["日期"] <= last - pd.DateOffset(months=12)]["日期"].max()].set_index("Wind代码")
P_aug = M[M["日期"] == pd.Timestamp("2026-08-31")].set_index("Wind代码")
snap = pd.DataFrame({"名称": S["名称"], "分组": S["研究分组"], "股息率(预期优先)%": S["股息率"],
                     "外资税后利差_US(pp)": S["利差_US"], "内资税后利差_CN(pp)": S["利差_CN"],
                     "港股通占港股%": S["港股通_占流通%"], "非内资占港股%": S["非内资_占流通%"], "中资占港股%": S["中资_占流通%"],
                     "港股通12M变化pp": S["港股通_占流通%"] - P12["港股通_占流通%"],
                     "非内资12M变化pp": S["非内资_占流通%"] - P12["非内资_占流通%"],
                     "非港股通12M变化pp": S["非港股通_占流通%"] - P12["非港股通_占流通%"],
                     "港股通自8月末pp": S["港股通_占流通%"] - P_aug["港股通_占流通%"],
                     "非内资自8月末pp": S["非内资_占流通%"] - P_aug["非内资_占流通%"]})
snap.sort_values(["分组", "外资税后利差_US(pp)"], ascending=[True, False]).to_csv(OUT / "38_当前个股快照.csv", encoding="utf-8-sig")

# ---------- 图 ----------
hi = M.groupby("日期")["US10Y"].first() >= 4
fig, axes = plt.subplots(3, 1, figsize=(13, 10), sharex=True)
for ax, k, t in zip(axes, ["港股通", "非内资", "非港股通"], ["港股通持股", "国际+本地中介持股（灰色=Wind归类迁移期）", "非港股通中介合计（国际+本地+中资）"]):
    for grp, c in zip(["A", "B", "C"], ["#08306b", "#e39b2d", "#9e9e9e"]):
        x = G[G["分组"] == grp].set_index("日期")[k]
        ax.plot(x.index, x, color=c, lw=1.5, label=GROUP_NAME[grp])
    ax.fill_between(hi.index, 0, 1, where=hi.values, transform=ax.get_xaxis_transform(), color="#d9534f", alpha=.12, lw=0)
    ax.set_ylabel("占港股上市市值 %")
    ax.set_title(t)
axes[0].legend(loc="upper left")
axes[1].axvspan(*RECLASS, color="#cccccc", alpha=.4)
fig.suptitle("指标股中介持股结构（红色阴影=美债10Y≥4%）")
fig.tight_layout()
fig.savefig(OUT / "fig6_分组中介持股.png", dpi=140)

fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), sharey=False)
for ax, grp in zip(axes, ["A", "B"]):
    x = F[F["研究分组"] == grp].set_index("日期")
    roll_sb = x["港股通_净买入亿"].rolling(6).sum()
    roll_fx = x["非港股通_净买入亿"].rolling(6).sum()
    ax.plot(roll_sb.index, roll_sb, color="#08306b", label="港股通")
    ax.plot(roll_fx.index, roll_fx, color="#e39b2d", label="非港股通中介合计")
    ax.fill_between(hi.index, 0, 1, where=hi.values, transform=ax.get_xaxis_transform(), color="#d9534f", alpha=.12, lw=0)
    ax.axhline(0, color="black", lw=.6)
    ax.set_title(GROUP_NAME[grp] + "：滚动6个月估算净买入（亿港元）")
    ax.legend(loc="upper left")
fig.tight_layout()
fig.savefig(OUT / "fig7_净买入_滚动6M.png", dpi=140)

fig, ax = plt.subplots(figsize=(9, 6))
s = snap[snap["分组"].isin(["A", "B"])]
for grp, c in [("A", "#08306b"), ("B", "#e39b2d")]:
    z = s[s["分组"] == grp]
    ax.scatter(z["外资税后利差_US(pp)"], z["非港股通12M变化pp"], color=c, label=GROUP_NAME[grp], s=40)
    for i, r in z.iterrows():
        ax.annotate(r["名称"], (r["外资税后利差_US(pp)"], r["非港股通12M变化pp"]), fontsize=8, xytext=(3, 3), textcoords="offset points")
ax.axhline(0, color="black", lw=.6)
ax.axvline(0, color="black", lw=.6)
ax.set_xlabel("外资视角税后股息率 − 美债10Y (pp)")
ax.set_ylabel("近12个月 非港股通中介 持股变化 (pp)")
ax.set_title(f"个股：对美债的税后利差 vs 非港股通中介持股变化（{last.date()}）")
ax.legend()
fig.tight_layout()
fig.savefig(OUT / "fig8_个股利差vs外资持股变化.png", dpi=140)
print("done", last.date())
