"""美联储加息/降息周期下港股红利指数表现分析。

输入: data/hk_dividend_fed_cycle_20261103.xlsx（Wind/Bloomberg/FRED 整理数据，不入库）
输出: output/ 下的 CSV 统计表与 PNG 图

运行: python3 analysis/fed_cycle_hk_dividend.py [xlsx路径]
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

# 研究对象（全收益口径）
IDX = {
    "HSHDYI": "恒生高股息率全收益(HKD)",
    "HSHYLV": "恒生港股通红利低波动全收益(HKD)",
    "CSI港股通高股息": "中证港股通高股息港元全收益(HKD)",
    "HSI": "恒生指数全收益(HKD)",
    "HSTECH": "恒生科技全收益(HKD)",
    "SP股息贵族(USD)": "标普500股息贵族全收益(USD)",
}
BENCH = "HSI"
DIV = ["HSHDYI", "HSHYLV", "CSI港股通高股息"]

xl = pd.read_excel(XLSX, sheet_name=None)
daily = xl["指数日度"].set_index("日期").sort_index()
px = daily[list(IDX.values())].rename(columns={v: k for k, v in IDX.items()})
cycles = xl["周期划分"].copy()
expl = xl["解释变量"].set_index("日期").sort_index()
val = xl["估值与利差"].set_index("日期").sort_index()
monthly = xl["月度全收益"].copy()
fedd = xl["联储日度利率"].set_index("日期").sort_index()

# HSHDYI 指数 2012-12-10 才正式发布，之前为回溯
HSHDYI_LAUNCH = pd.Timestamp("2012-12-10")
LAST = px.index.max()


def seg_stats(s: pd.Series, start, end):
    """区间内 [start, end] 的累计收益/年化/波动/最大回撤；基点取 start 之前最后一个交易日。"""
    s = s.dropna()
    before = s[s.index < start]
    inside = s[(s.index >= start) & (s.index <= end)]
    if inside.empty:
        return dict(累计=np.nan, 年化=np.nan, 波动=np.nan, 最大回撤=np.nan, 实际起点=None)
    if before.empty and (inside.index[0] - start).days > 120:
        return dict(累计=np.nan, 年化=np.nan, 波动=np.nan, 最大回撤=np.nan, 实际起点=None)
    if before.empty and (inside.index[0] - start).days > 7:
        # 序列在区段开始后不久才有数据（如恒指全收益始于2004-10-11）：用首个可得点，并标注实际起点
        path = inside
    else:
        path = pd.concat([before.iloc[-1:], inside]) if not before.empty else inside
    tot = path.iloc[-1] / path.iloc[0] - 1
    yrs = (path.index[-1] - path.index[0]).days / 365.25
    ann = (1 + tot) ** (1 / yrs) - 1 if yrs > 0.25 else np.nan
    r = path.pct_change().dropna()
    vol = r.std() * np.sqrt(250) if len(r) > 20 else np.nan
    mdd = (path / path.cummax() - 1).min()
    return dict(累计=tot, 年化=ann, 波动=vol, 最大回撤=mdd, 实际起点=path.index[0].date())


# ---------- 1. 周期分段：调整段 + 调整后维持段 + 整个周期 ----------
cycles["段起"] = pd.to_datetime(cycles["FRED首次调整记录日"])
cycles["段止"] = pd.to_datetime(cycles["FRED末次调整记录日"])
cycles["下一阶段"] = pd.to_datetime(cycles["下一阶段FRED记录日"])
segs = []
for _, c in cycles.iterrows():
    nxt = c["下一阶段"] if pd.notna(c["下一阶段"]) else LAST + pd.Timedelta(days=1)
    lab = c["降息类型（研究标签）"] if c["方向"] == "降息" else "加息"
    segs.append((c["周期编号"], c["方向"], lab, "调整段", c["段起"], c["段止"]))
    if c["段止"] + pd.Timedelta(days=1) < nxt - pd.Timedelta(days=1):
        segs.append((c["周期编号"], c["方向"], lab, "维持段", c["段止"] + pd.Timedelta(days=1), nxt - pd.Timedelta(days=1)))
    segs.append((c["周期编号"], c["方向"], lab, "整个周期(至反向首调前)", c["段起"], nxt - pd.Timedelta(days=1)))

rows = []
for cid, d, lab, kind, a, b in segs:
    b = min(b, LAST)
    base = seg_stats(px[BENCH], a, b)
    for k in IDX:
        st = seg_stats(px[k], a, b)
        rows.append(dict(周期=cid, 方向=d, 标签=lab, 区段=kind, 起=a.date(), 止=b.date(), 指数=k, **st,
                         超额HSI_累计=st["累计"] - base["累计"] if k != BENCH else np.nan,
                         超额HSI_年化=st["年化"] - base["年化"] if k != BENCH else np.nan,
                         回溯数据=(k == "HSHDYI" and a < HSHDYI_LAUNCH)))
seg_df = pd.DataFrame(rows)
seg_df.to_csv(OUT / "01_周期分段收益.csv", index=False, encoding="utf-8-sig")

# ---------- 2. 首次加/降息事件窗口 ----------
HORIZ = [-60, -20, 20, 60, 120, 250]
ev_rows = []
for _, c in cycles.iterrows():
    dec = pd.to_datetime(c["首次决议日"])
    for k in IDX:
        s = px[k].dropna()  # 每个序列用自己的交易日历（港股假期不同于美股）
        rec = dict(周期=c["周期编号"], 方向=c["方向"], 标签=c["降息类型（研究标签）"] if c["方向"] == "降息" else "加息",
                   指数=k)
        t0_pos = s.index.searchsorted(dec, side="right")  # 决议在美东下午，港股 T0 = 决议日后首个交易日
        base_pos = t0_pos - 1  # T-1 收盘为基点
        rec["T0"] = s.index[t0_pos].date() if t0_pos < len(s) else None
        for h in HORIZ:
            p0, p1 = (base_pos + h, base_pos) if h < 0 else (base_pos, base_pos + h)
            ok = 0 <= p0 and p1 < len(s) and t0_pos < len(s) and base_pos >= 0 \
                and (s.index[t0_pos] - dec).days <= 7  # 序列须在决议前已存在
            rec[f"T{h:+d}"] = s.iloc[p1] / s.iloc[p0] - 1 if ok else np.nan
        ev_rows.append(rec)
ev = pd.DataFrame(ev_rows)
evx = ev.copy()
hcols = [f"T{h:+d}" for h in HORIZ]
bench_ev = ev[ev["指数"] == BENCH].set_index("周期")[hcols]
for cid in evx["周期"].unique():
    m = evx["周期"] == cid
    evx.loc[m, hcols] = evx.loc[m, hcols].values - bench_ev.loc[cid].values
ev.to_csv(OUT / "02_首次调整事件窗口_绝对收益.csv", index=False, encoding="utf-8-sig")
evx[evx["指数"] != BENCH].to_csv(OUT / "03_首次调整事件窗口_超额HSI.csv", index=False, encoding="utf-8-sig")

# ---------- 3. 按日度阶段标签汇总（加息段/加息后维持/降息段/降息后维持） ----------
phase_all = fedd["阶段（按FRED记录日）"]
dr = {k: px[k].dropna().pct_change().dropna() for k in IDX}
ph_rows = []
for k in IDX:
    for p, g in dr[k].groupby(phase_all.reindex(dr[k].index, method="ffill")):
        g = g.dropna()
        if k == "HSHDYI":
            g = g[g.index >= HSHDYI_LAUNCH]  # 只用发布后实盘
        if len(g) < 60:
            continue
        exc = (g - dr[BENCH].reindex(g.index)).dropna()  # 美元资产与港股收盘时差，超额仅作参考
        ph_rows.append(dict(指数=k, 阶段=p, 交易日=len(g),
                            年化收益=(1 + g).prod() ** (250 / len(g)) - 1,
                            年化波动=g.std() * np.sqrt(250),
                            年化超额HSI=exc.mean() * 250 if k != BENCH else np.nan,
                            超额日胜率=(exc > 0).mean() if k != BENCH else np.nan))
ph = pd.DataFrame(ph_rows)
ph.to_csv(OUT / "04_按利率阶段年化(HSHDYI仅发布后).csv", index=False, encoding="utf-8-sig")

# ---------- 4. 月度回归：红利超额 ~ 利率/美元/中债/南向 ----------
m = monthly.set_index("月份")
m = m[m["月份状态"] == "完整月"]
y = m["恒生高股息率全收益超额恒指收益"].rename("HSHDYI超额")
em = expl.resample("MS").last()
X = pd.DataFrame({
    "ΔUS10Y(pp)": em["美国10年期国债恒定期限收益率"].diff(),
    "ΔUS实际10Y(pp)": em["美国10年期实际国债恒定期限收益率"].diff(),
    "DXY变动": em["美元指数(USD)"].pct_change(),
    "ΔCN10Y(pp)": em["中国10年期国债Wind重取对照"].diff(),
    "南向净买入(千亿HKD)": expl["南向合计每日净买入"].resample("MS").sum(min_count=1) / 1000,
})
fed_m = fedd["统一中点(%)"].resample("MS").last()
X["ΔFFR(pp)"] = fed_m.diff()
reg_rows = []
for name, cols, start in [
    ("全样本 2007-2026（含回溯）", ["ΔUS10Y(pp)", "DXY变动", "ΔFFR(pp)"], "2007-08-01"),
    ("发布后 2013-2026", ["ΔUS10Y(pp)", "DXY变动", "ΔFFR(pp)"], "2013-01-01"),
    ("发布后+中债", ["ΔUS10Y(pp)", "ΔCN10Y(pp)", "DXY变动"], "2013-01-01"),
    ("实际利率", ["ΔUS实际10Y(pp)", "ΔCN10Y(pp)", "DXY变动"], "2013-01-01"),
    ("含南向资金 2015+", ["ΔUS10Y(pp)", "ΔCN10Y(pp)", "DXY变动", "南向净买入(千亿HKD)"], "2015-01-01"),
]:
    dfr = pd.concat([y, X[cols]], axis=1).loc[start:].dropna()
    res = sm.OLS(dfr["HSHDYI超额"], sm.add_constant(dfr[cols])).fit(cov_type="HAC", cov_kwds={"maxlags": 3})
    for v in ["const"] + cols:
        reg_rows.append(dict(模型=name, 变量=v, 系数=res.params[v], t值=res.tvalues[v], p值=res.pvalues[v],
                             N=int(res.nobs), R2=res.rsquared))
reg = pd.DataFrame(reg_rows)
reg.to_csv(OUT / "05_月度回归_HSHDYI超额.csv", index=False, encoding="utf-8-sig")

# ---------- 5. 股息率利差：各周期起点水平 & 之后12个月超额 ----------
sp = val[["HSHDYI近12个月股息率", "HSHDYI股息率减美国10Y(bp)", "HSHDYI股息率减中国10Y Wind重取(bp)",
          "HSHDYI滚动市盈率", "HSHDYI最新财报市净率"]].dropna(how="all")
pct = lambda s, v: (s.dropna() <= v).mean()
sp_rows = []
for _, c in cycles.iterrows():
    a = c["段起"]
    if a < sp.index.min():
        continue
    row = sp.loc[:a].iloc[-1]
    hist = sp.loc[:a]
    rec = dict(周期=c["周期编号"], 方向=c["方向"], 日期=a.date())
    for col in sp.columns:
        rec[col] = row[col]
        rec[col + "_全样本分位"] = pct(sp[col], row[col])
    sp_rows.append(rec)
cur = sp.ffill().iloc[-1]  # FRED 美债最新到 10-30，利差沿用最近值
rec = dict(周期="当前", 方向="", 日期=sp.index[-1].date())
for col in sp.columns:
    rec[col] = cur[col]
    rec[col + "_全样本分位"] = pct(sp[col], cur[col])
sp_rows.append(rec)
spd = pd.DataFrame(sp_rows)
spd.to_csv(OUT / "06_周期起点估值与利差.csv", index=False, encoding="utf-8-sig")

# 利差分组 -> 未来12个月超额（月度，重叠样本，仅作描述）
ms = sp.resample("MS").last()
lvl = m["恒生高股息率全收益月末值"] / m["恒生指数全收益月末值"]
fwd = (lvl.shift(-12) / lvl - 1).rename("未来12M超额")
q = pd.concat([ms["HSHDYI股息率减美国10Y(bp)"], fwd.reindex(ms.index)], axis=1).dropna()
q["利差五分位"] = pd.qcut(q["HSHDYI股息率减美国10Y(bp)"], 5, labels=["Q1低", "Q2", "Q3", "Q4", "Q5高"])
qq = q.groupby("利差五分位", observed=True).agg(利差均值=("HSHDYI股息率减美国10Y(bp)", "mean"),
                                           未来12M超额均值=("未来12M超额", "mean"),
                                           胜率=("未来12M超额", lambda s: (s > 0).mean()), 月数=("未来12M超额", "size"))
qq.to_csv(OUT / "07_美债利差分组_未来12M超额.csv", encoding="utf-8-sig")

# ---------- 6. 图 ----------
fig, ax = plt.subplots(2, 1, figsize=(13, 8.5), sharex=True, gridspec_kw={"height_ratios": [2.2, 1]})
rel = (px["HSHDYI"] / px["HSI"]).dropna()
rel = rel / rel.iloc[0]
for _, c in cycles.iterrows():
    nxt = c["下一阶段"] if pd.notna(c["下一阶段"]) else LAST
    col = "#d9534f" if c["方向"] == "加息" else "#5b9bd5"
    for a_ in ax:
        a_.axvspan(c["段起"], c["段止"], color=col, alpha=0.18, lw=0)
ax[0].plot(rel.index, rel, color="#1f3b57", lw=1.4, label="HSHDYI全收益 / 恒指全收益")
ax[0].axvline(HSHDYI_LAUNCH, color="gray", ls="--", lw=0.8)
ax[0].text(HSHDYI_LAUNCH, rel.max() * 0.95, " 指数发布(此前为回溯)", color="gray", fontsize=9)
ax[0].set_ylabel("相对强弱（2007-06=1）")
ax[0].set_title("港股高股息相对恒指表现与美联储周期（红=加息调整段，蓝=降息调整段）")
ax[0].legend(loc="upper left")
ax[1].step(fedd.index, fedd["统一上界(%)"], where="post", color="#333", lw=1.2, label="联邦基金目标上限")
ax[1].plot(expl.index, expl["美国10年期国债恒定期限收益率"], color="#e39b2d", lw=0.9, label="美债10Y")
ax[1].set_ylabel("%")
ax[1].legend(loc="upper right")
ax[1].set_xlim(pd.Timestamp("2007-01-01"), LAST)
fig.tight_layout()
fig.savefig(OUT / "fig1_相对强弱与周期.png", dpi=140)

fig, ax = plt.subplots(figsize=(11, 5))
show = evx[(evx["指数"] == "HSHDYI")].set_index("周期")
cyc_lab = {r["周期编号"]: f'{r["周期编号"]} {r["方向"]} {pd.to_datetime(r["首次决议日"]).year}' for _, r in cycles.iterrows()}
show = show[["T+60", "T+120", "T+250"]].dropna(how="all")
show.index = [cyc_lab[i] for i in show.index]
show.mul(100).plot.bar(ax=ax, color=["#9ecae1", "#4292c6", "#08306b"])
ax.axhline(0, color="black", lw=0.8)
ax.set_ylabel("HSHDYI相对恒指超额(%)")
ax.set_title("首次加/降息后HSHDYI相对恒指超额收益（T-1收盘为基点，交易日）")
plt.xticks(rotation=0)
fig.tight_layout()
fig.savefig(OUT / "fig2_首次调整后超额.png", dpi=140)

fig, ax = plt.subplots(figsize=(13, 4.5))
ax.plot(sp.index, sp["HSHDYI股息率减美国10Y(bp)"], label="HSHDYI股息率 − 美债10Y", color="#1f3b57")
ax.plot(sp.index, sp["HSHDYI股息率减中国10Y Wind重取(bp)"], label="HSHDYI股息率 − 中债10Y", color="#e39b2d")
for _, c in cycles.iterrows():
    if c["段止"] >= sp.index.min():
        ax.axvspan(max(c["段起"], sp.index.min()), c["段止"], color="#d9534f" if c["方向"] == "加息" else "#5b9bd5", alpha=0.18, lw=0)
ax.axhline(0, color="black", lw=0.6)
ax.set_ylabel("bp")
ax.legend()
ax.set_title("高股息股息率利差")
fig.tight_layout()
fig.savefig(OUT / "fig3_股息率利差.png", dpi=140)

print("done ->", OUT)
