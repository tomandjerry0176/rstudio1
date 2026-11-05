"""用 Bloomberg 个股一致预期核对指数层面的 2027E 增速差异（默认恒生科技）。

背景：本套 Wind 自下而上（指数权重法）恒生科技 2027E 增速 +50.1%，BBG 指数 FY2/FY1 为 +40.7%。
可能原因：(1) BBG 对互联网公司默认经调整口径（BEST_EPS），Wind 为归母净利润；
          (2) BBG 指数 FY1/FY2 为“相对财年”，3 月年结公司的 FY1 可能已是截至 2027-03 的财年，本套为日历年。
做法：同一组权重、同一公式（Σ权重 × EPS/股价），分别用 BBG 经调整 EPS、BBG GAAP EPS 和本套 Wind 利润计算增速，
     并逐只列出差异与 BBG 的财年截止日。

依赖：Bloomberg 终端 + xbbg（pip install xbbg）、pandas、openpyxl；以及港股脚本输出的 Excel（个股明细）。
用法：
    python bbg_check_index.py --wind 港股_恒生一级_2027盈利预期_2026-11-04.xlsx
    python bbg_check_index.py --index "HSI Index" --wind ...      # 其他指数同理
字段名需在终端 FLDS 中核对：BEST_EPS（BEst 标准口径，多为经调整）、BEST_EPS_GAAP（GAAP 口径）、
BEST_PERIOD_END_DT（该财年截止日）；货币用 EQY_FUND_CRNCY=HKD 覆盖，与港元股价一致。
"""
import argparse
import re
from pathlib import Path

import numpy as np
import pandas as pd

FIELDS = {"BBG经调整": "BEST_EPS", "BBG_GAAP": "BEST_EPS_GAAP"}
PERIODS = {"FY1": "1FY", "FY2": "2FY"}
CURRENCY = "HKD"


def bbg_to_wind(ticker):
    """'700 HK Equity' / '700 HK' -> '0700.HK'"""
    m = re.match(r"\s*(\d+)\s+HK", str(ticker))
    return f"{int(m.group(1)):04d}.HK" if m else None


def pick(frame, *names):
    cols = {c.lower().replace(" ", "_"): c for c in frame.columns}
    for n in names:
        if n.lower() in cols:
            return frame[cols[n.lower()]]
    raise KeyError(f"BBG 返回缺少列 {names}，实际列：{list(frame.columns)}")


def fetch_bbg(blp, index):
    members = blp.bds(index, "INDX_MWEIGHT")
    if members is None or members.empty:
        raise RuntimeError(f"{index} INDX_MWEIGHT 为空：检查指数代码与终端登录")
    tick = pick(members, "member_ticker_and_exchange_code", "index_member")
    wt = pd.to_numeric(pick(members, "percentage_weight", "weight"), errors="coerce")
    tickers = [f"{t} Equity" if not str(t).endswith("Equity") else str(t) for t in tick]
    d = pd.DataFrame({"bbg": tickers, "权重": (wt / wt.sum()).values})
    d["股票代码"] = d["bbg"].map(bbg_to_wind)

    base = blp.bdp(d["bbg"].tolist(), ["NAME", "PX_LAST", "CRNCY"])
    base.index = base.index.astype(str)
    d["名称"] = d["bbg"].map(pick(base, "name"))
    d["股价"] = d["bbg"].map(pd.to_numeric(pick(base, "px_last"), errors="coerce"))
    d["股价币种"] = d["bbg"].map(pick(base, "crncy"))
    for pname, pcode in PERIODS.items():
        got = blp.bdp(d["bbg"].tolist(), list(FIELDS.values()) + ["BEST_PERIOD_END_DT"],
                      BEST_FPERIOD_OVERRIDE=pcode, EQY_FUND_CRNCY=CURRENCY)
        got.index = got.index.astype(str)
        for label, fld in FIELDS.items():
            d[f"{label}_EPS_{pname}"] = d["bbg"].map(pd.to_numeric(pick(got, fld), errors="coerce"))
        d[f"财年截止_{pname}"] = d["bbg"].map(pick(got, "best_period_end_dt")).astype(str).str[:10]
    return d


def weighted_growth(w, ey0, ey1):
    ok = w.notna() & ey0.notna() & ey1.notna()
    e0, e1 = (w[ok] * ey0[ok]).sum(), (w[ok] * ey1[ok]).sum()
    contrib = (w * (ey1 - ey0) / e0).where(ok) if e0 > 0 else pd.Series(np.nan, index=w.index)
    return (e1 / e0 - 1 if e0 > 0 else np.nan), contrib, float(w[ok].sum())


def main():
    ap = argparse.ArgumentParser(description="BBG 个股一致预期核对指数 2027E 增速")
    ap.add_argument("--index", default="HSTECH Index")
    ap.add_argument("--wind", required=True, help="港股脚本输出的 Excel（读取个股明细）")
    ap.add_argument("--out", help="输出 Excel，默认与 --wind 同目录")
    args = ap.parse_args()

    from xbbg import blp
    d = fetch_bbg(blp, args.index)

    wind = pd.read_excel(args.wind, sheet_name="个股明细")
    wind["股票代码"] = wind["股票代码"].astype(str).str.upper()
    wind = wind.set_index("股票代码")
    mv = wind["MKT_CAP_ARD"].where(wind["MKT_CAP_ARD"] > 0)
    d["Wind_EY_2026"] = d["股票代码"].map(wind["NP2026"] / mv)
    d["Wind_EY_2027"] = d["股票代码"].map(wind["NP2027"] / mv)
    d["Wind口径"] = d["股票代码"].map(wind.get("口径"))
    d["Wind_FY2026来源"] = d["股票代码"].map(wind.get("FY2026来源"))

    results = []
    g, c, cov = weighted_growth(d["权重"], d["Wind_EY_2026"], d["Wind_EY_2027"])
    d["贡献_Wind(百分点)"] = c * 100
    results.append(("Wind 归母净利润（日历年）", g, cov))
    for label in FIELDS:
        ey1 = d[f"{label}_EPS_FY1"] / d["股价"]
        ey2 = d[f"{label}_EPS_FY2"] / d["股价"]
        g, c, cov = weighted_growth(d["权重"], ey1, ey2)
        d[f"贡献_{label}(百分点)"] = c * 100
        results.append((f"{label} EPS（FY2/FY1，相对财年）", g, cov))
    d["贡献差_Wind减BBG经调整"] = d["贡献_Wind(百分点)"] - d["贡献_BBG经调整(百分点)"]

    summary = pd.DataFrame(results, columns=["口径", "增速(权重法)", "权重覆盖"])
    print(f"== {args.index}：同一组权重、同一公式下的 FY2/FY1 增速")
    print(summary.to_string(index=False, formatters={"增速(权重法)": "{:+.1%}".format, "权重覆盖": "{:.0%}".format}))
    cols = ["名称", "股票代码", "权重", "财年截止_FY1", "财年截止_FY2", "Wind口径", "贡献_Wind(百分点)",
            "贡献_BBG经调整(百分点)", "贡献_BBG_GAAP(百分点)", "贡献差_Wind减BBG经调整"]
    diff = d.reindex(d["贡献差_Wind减BBG经调整"].abs().sort_values(ascending=False).index)
    print("\n== 贡献差最大的 10 只（百分点）")
    print(diff[cols].head(10).round(2).to_string(index=False))
    off = d[~d["财年截止_FY1"].str.startswith("2026-12")]
    if len(off):
        print(f"\n== BBG FY1 不是截至 2026-12 的公司（{len(off)} 只，按相对财年汇总）")
        print(off[["名称", "权重", "财年截止_FY1", "财年截止_FY2"]].round(4).to_string(index=False))

    out = Path(args.out) if args.out else Path(args.wind).with_name(
        f"BBG核对_{args.index.split()[0]}_{Path(args.wind).stem.split('_')[-1]}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="增速对照", index=False)
        diff.to_excel(xw, sheet_name="个股对照", index=False)
    print(f"\n已保存：{out}")


if __name__ == "__main__":
    main()
