"""高利率环境下港股红利表现（主框架）；美联储加/降息阶段作为辅助切分。

高利率定义（三种口径并列，避免结论依赖单一阈值）：
  A. 美债10Y ≥ 4%
  B. 联邦基金目标上限 ≥ 4%
  C. 美债10Y实际收益率 ≥ 1.5%
运行: python3 analysis/high_rate_regime.py [xlsx路径]
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "hk_dividend_fed_cycle_20261103.xlsx"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)

xl = pd.read_excel(XLSX, sheet_name=["指数日度", "解释变量", "联储日度利率"])
d = xl["指数日度"].set_index("日期").sort_index()
ex = xl["解释变量"].set_index("日期").sort_index()
fed = xl["联储日度利率"].set_index("日期").sort_index()

IDX = {"HSHDYI": "恒生高股息率全收益(HKD)", "HSHYLV": "恒生港股通红利低波动全收益(HKD)",
       "CSI港股通高股息": "中证港股通高股息港元全收益(HKD)", "HSI": "恒生指数全收益(HKD)",
       "HSTECH": "恒生科技全收益(HKD)"}
cal = d["恒生指数全收益(HKD)"].dropna().index  # 港股交易日历
cal = cal[cal >= "2007-06-29"]
px = d[list(IDX.values())].rename(columns={v: k for k, v in IDX.items()}).reindex(cal)

st = pd.DataFrame(index=cal)
st["US10Y"] = ex["美国10年期国债恒定期限收益率"].reindex(cal, method="ffill")
st["Real10Y"] = ex["美国10年期实际国债恒定期限收益率"].reindex(cal, method="ffill")
st["FFR上限"] = fed["统一上界(%)"].reindex(cal, method="ffill")
st["联储阶段"] = fed["阶段（按FRED记录日）"].reindex(cal, method="ffill").str.replace("（含段内暂停）", "", regex=False)
REG = {"A_美债10Y≥4%": st["US10Y"] >= 4, "B_联邦基金≥4%": st["FFR上限"] >= 4, "C_实际利率≥1.5%": st["Real10Y"] >= 1.5}

ret = px.pct_change(fill_method=None)


def stats(mask):
    out = {"交易日": int(mask.sum())}
    for k in IDX:
        r = ret.loc[mask, k].dropna()
        if len(r) < 40:
            out[k + "_年化"] = np.nan
            continue
        out[k + "_年化"] = (1 + r).prod() ** (250 / len(r)) - 1
        if k != "HSI":
            e = (r - ret.loc[r.index, "HSI"]).dropna()
            out[k + "_年化超额"] = e.mean() * 250
    return out


# 1. 高利率 vs 非高利率
rows = []
for name, m in REG.items():
    for lab, mm in [("高利率", m), ("非高利率", ~m)]:
        rows.append(dict(口径=name, 环境=lab, **stats(mm)))
t1 = pd.DataFrame(rows)
t1.to_csv(OUT / "20_高利率vs非高利率_年化.csv", index=False, encoding="utf-8-sig")

# 2. 高利率内部按联储阶段（辅助）
rows = []
for name, m in REG.items():
    for ph in st["联储阶段"].dropna().unique():
        mm = m & (st["联储阶段"] == ph)
        if mm.sum() >= 40:
            rows.append(dict(口径=name, 联储阶段=ph, **stats(mm)))
pd.DataFrame(rows).to_csv(OUT / "21_高利率x联储阶段_年化.csv", index=False, encoding="utf-8-sig")

# 3. 高利率区间（口径A，合并间隔<20个交易日的断点，保留≥60个交易日的区间）
m = REG["A_美债10Y≥4%"].astype(int)
grp = (m.diff().fillna(0) != 0).cumsum()
spells = [(g.index[0], g.index[-1]) for _, g in m.groupby(grp) if g.iloc[0] == 1]
merged = []
for a, b in spells:
    if merged and cal.get_loc(a) - cal.get_loc(merged[-1][1]) < 20:
        merged[-1] = (merged[-1][0], b)
    else:
        merged.append((a, b))
rows = []
for a, b in merged:
    n = cal.get_loc(b) - cal.get_loc(a)
    if n < 60:
        continue
    p0 = cal[max(cal.get_loc(a) - 1, 0)]
    rec = dict(起=a.date(), 止=b.date(), 交易日=n, 美债10Y均值=st.loc[a:b, "US10Y"].mean(),
               联邦基金上限区间=f'{st.loc[a:b, "FFR上限"].min():.2f}-{st.loc[a:b, "FFR上限"].max():.2f}')
    for k in IDX:
        s = px[k]
        rec[k] = s.loc[b] / s.loc[p0] - 1 if pd.notna(s.loc[p0]) and pd.notna(s.loc[b]) else np.nan
    for k in IDX:
        if k != "HSI":
            rec[k + "_超额"] = rec[k] - rec["HSI"]
    rows.append(rec)
pd.DataFrame(rows).to_csv(OUT / "22_高利率区间明细_美债4%.csv", index=False, encoding="utf-8-sig")
print("done")
