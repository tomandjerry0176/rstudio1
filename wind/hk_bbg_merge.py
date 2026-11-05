"""把 BBG 取数模板（刷新后发回的 Excel）合并进港股流程。

用法：
    python hk_bbg_merge.py --wind 港股_恒生一级_2027盈利预期_<日期>.xlsx --bbg BBG取数模板_港股恒生综合.xlsx --mode fill
    python hk_bbg_merge.py --wind ... --bbg ... --mode full
  fill：Wind 为主，只对 Wind 无 2027E 预测（预测来源=无预测）的公司用 BBG 补；同一公司的多条股票线只补一条
  full：全部改用 BBG 一致预期（市值、行业、币种折算后的已报基期沿用 Wind 明细），另出 Wind 与 BBG 逐行业对照
  full-fy：全部用 BBG 相对财年 FY1 → FY2（不做日历化），与 BBG 指数 FY1/FY2 的构造一致，用于拆解指数层面的差异
输出与港股脚本同格式（行业汇总 / 个股明细 / 统计说明 + 看板），可直接作为 index_consensus.py 的 --h 输入。

BBG 相对财年 → 日历年：1FY 的截止日 = 财年截止月 m（Wind 明细“财年截止月”，缺失时用 BBG EQY_FISCAL_YR_END，再缺按 12 月）
  下、截止日 + 3 个月晚于数据日期的第一个财年（即尚未报出的第一个财年）；模板 END_1FY 列（BEST_PERIOD_END_DATE，1FY）
  有值时优先使用，与推算不一致的公司会打印出来。
  1FY/2FY/3FY 归到截止年份 Y 后，
  CY(Y) = m/12 × V(Y) + (1 − m/12) × V(Y+1)；V(Y) 缺 BBG 时（已报财年）用 Wind 明细中的已报值（已折港元）。
历史对比日：若某财年（截止于当前 1FY 前一年）在对比日之后约 3 个月内才报出，则对比日的 1FY/2FY 相当于当前的
  前一财年/1FY，需错开一位（主要影响 6 月年结公司）；按截止日 + 3 个月近似判断，统计说明中注明。
"""
import argparse
import re
from pathlib import Path

import sys

import numpy as np
import pandas as pd

# 与 hk_consensus_2027.py 放在同一目录（financials_lib）运行；共享模块 sw1_consensus_2027.py 在财报输出目录
CORE_DIR = Path("/Users/amber/Documents/本地数据库/财报")
try:
    import sw1_consensus_2027 as core
except ImportError:
    sys.path.append(str(CORE_DIR))
    import sw1_consensus_2027 as core

MILLION = 1e6
REPORT_LAG_MONTHS = 3
# 同一公司的多条股票线：key 不补，以 value 那条为准（避免利润重复计算）
DUPLICATE_LINES = {"0087.HK": "0019.HK"}   # 太古股份公司 B → 以 A 为准


def read_bbg(path):
    raw = pd.read_excel(path, sheet_name="BBG取数", header=3)
    raw = raw[raw["Wind代码"].notna()].copy()
    raw["Wind代码"] = raw["Wind代码"].astype(str).str.upper()
    raw = raw.set_index("Wind代码")
    params = pd.read_excel(path, sheet_name="参数", header=None).dropna(subset=[0])
    windows = {str(r[0]).replace("前", ""): pd.Timestamp(r[1]) for r in params.itertuples(index=False)
               if str(r[0]).endswith("前")}
    date = pd.Timestamp(params[params[0] == "数据日期"].iloc[0, 1])

    def num(col):
        return pd.to_numeric(raw[col], errors="coerce") if col in raw else pd.Series(np.nan, index=raw.index)

    def dt(col):
        return pd.to_datetime(raw[col], errors="coerce") if col in raw else pd.Series(pd.NaT, index=raw.index)

    return raw, num, dt, windows, date


def to_calendar(values_by_year, month, year):
    """values_by_year: {截止年份: 值}；month: 截止月。返回日历年 year 的值（缺任一所需财年则为 NaN）。"""
    if pd.isna(month):
        return np.nan
    if int(month) == 12:
        return values_by_year.get(year, np.nan)
    a, b = values_by_year.get(year, np.nan), values_by_year.get(year + 1, np.nan)
    return month / 12 * a + (1 - month / 12) * b if pd.notna(a) and pd.notna(b) else np.nan


def fye_month(v):
    """BBG EQY_FISCAL_YR_END（'03/31'、'3/31'、日期等）→ 月份。"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if hasattr(v, "month"):
        return float(v.month)
    m = re.match(r"\s*(\d{1,2})[/-]", str(v))
    return float(m.group(1)) if m and 1 <= int(m.group(1)) <= 12 else np.nan


def first_unreported_end(month, date):
    """截止月为 month 的财年中，截止日 + REPORT_LAG_MONTHS 晚于 date 的第一个财年的截止日。"""
    end = pd.Timestamp(date.year - 2, int(month), 1) + pd.offsets.MonthEnd(0)
    while end + pd.DateOffset(months=REPORT_LAG_MONTHS) <= date:
        end = pd.Timestamp(end.year + 1, int(month), 1) + pd.offsets.MonthEnd(0)
    return end


def bbg_calendar(wind, raw, num, dt, windows, date, by, fy):
    """返回 DataFrame：BBG_NP{by}、BBG_NP{fy}、BBG_NP{fy}_{w}前、BBG财年截止月、BBG口径。金额为港元（元）。"""
    out = pd.DataFrame(index=wind.index)
    end_given = dt("END_1FY").reindex(wind.index)
    bbg_fye = (raw["FYE(BBG)"].map(fye_month) if "FYE(BBG)" in raw else pd.Series(np.nan, index=raw.index)).reindex(wind.index)
    months = wind["财年截止月"].where(wind["财年截止月"].notna(), bbg_fye).fillna(12) if "财年截止月" in wind else bbg_fye.fillna(12)
    cur = {k: num(f"NI_{k}").reindex(wind.index) * MILLION for k in ("1FY", "2FY", "3FY")}
    past = {w: {k: num(f"NI_{k}_{w}前").reindex(wind.index) * MILLION for k in ("1FY", "2FY", "3FY")}
            for w in windows}
    # Wind 已报（港元）：FY{by}来源=已报 时的 FY{by}原币 × 折算系数
    reported = pd.Series(np.nan, index=wind.index)
    if f"FY{by}来源" in wind:
        rep = wind[f"FY{by}来源"].eq("已报")
        reported[rep] = (wind.loc[rep, f"FY{by}原币"] * wind.loc[rep, "折算系数"]).astype(float)

    inferred = pd.Series({c: first_unreported_end(months[c], date) for c in wind.index})
    diff = end_given.notna() & (end_given.dt.to_period("M") != inferred.dt.to_period("M"))
    if diff.any():
        print(f"提示：{int(diff.sum())} 只 BBG 1FY 截止日与推算不同，以 BBG 为准："
              + "、".join(f"{c}({end_given[c].date()} vs {inferred[c].date()})" for c in diff[diff].index[:10]), flush=True)
    rows = []
    for c in wind.index:
        if pd.isna(cur["1FY"].get(c)) and pd.isna(cur["2FY"].get(c)):
            rows.append({})
            continue
        e1 = end_given.get(c)
        if pd.isna(e1):
            e1 = inferred[c]
        m, y1 = e1.month, e1.year
        vals = {y1: cur["1FY"][c], y1 + 1: cur["2FY"][c], y1 + 2: cur["3FY"][c]}
        if pd.notna(reported.get(c)) and (y1 - 1) not in vals:
            vals[y1 - 1] = reported[c]
        rec = {"BBG财年截止月": m, f"BBG_NP{by}": to_calendar(vals, m, by),
               f"BBG_NP{fy}": to_calendar(vals, m, fy)}
        prev_end = e1 - pd.DateOffset(years=1)
        for w, t in windows.items():
            shift = t < prev_end + pd.DateOffset(months=REPORT_LAG_MONTHS)   # 对比日时前一财年尚未报出
            base_y = y1 - 1 if shift else y1
            pv = {base_y: past[w]["1FY"][c], base_y + 1: past[w]["2FY"][c], base_y + 2: past[w]["3FY"][c]}
            rec[f"BBG_NP{fy}_{w}前"] = to_calendar(pv, m, fy)
        rec["BBG口径"] = "12月财年" if m == 12 else "日历化"
        # 相对财年（不日历化）：FY1 → FY2；对比日若尚未滚动，则当时的 2FY 对应现在的 1FY，3FY 对应 2FY
        rec[f"BBGFY_NP{by}"], rec[f"BBGFY_NP{fy}"] = cur["1FY"][c], cur["2FY"][c]
        for w, t in windows.items():
            shift = t < prev_end + pd.DateOffset(months=REPORT_LAG_MONTHS)
            rec[f"BBGFY_NP{fy}_{w}前"] = past[w]["3FY" if shift else "2FY"][c]
        rows.append(rec)
    return pd.DataFrame(rows, index=wind.index)


def main():
    ap = argparse.ArgumentParser(description="BBG 一致预期合并进港股流程")
    ap.add_argument("--wind", required=True)
    ap.add_argument("--bbg", required=True)
    ap.add_argument("--mode", choices=["fill", "full", "full-fy"], default="fill")
    ap.add_argument("--out-dir")
    a = ap.parse_args()

    sheets = pd.read_excel(a.wind, sheet_name=None)
    wind = sheets["个股明细"].copy()
    wind["股票代码"] = wind["股票代码"].astype(str).str.upper()
    wind = wind.set_index("股票代码")
    years = sorted(int(c[2:]) for c in wind.columns if re.fullmatch(r"NP\d{4}", c))
    by, fy = years[0], years[-1]
    wins = [m.group(1) for c in wind.columns if (m := re.match(rf"NP{fy}_(.+)前$", c))]
    core.BASE_YEAR, core.FORECAST_YEAR = by, fy
    core.REVISION_WINDOWS = {w: None for w in wins}
    core.DATE = str(wind["请求数据日期"].dropna().iloc[0])[:10]
    core.OUT_DIR = Path(a.out_dir) if a.out_dir else Path(a.wind).parent

    raw, num, dt, windows, date = read_bbg(a.bbg)
    if str(date.date()) != core.DATE:
        print(f"警告：BBG 模板数据日期 {date.date()} 与 Wind {core.DATE} 不同", flush=True)
    missing_w = [w for w in wins if w not in windows]
    if missing_w:
        raise SystemExit(f"BBG 模板缺少对比日期：{missing_w}")
    b = bbg_calendar(wind, raw, num, dt, {w: windows[w] for w in wins}, date, by, fy)
    has_bbg = b[f"BBG_NP{fy}"].notna()
    print(f"BBG 有 {fy}E 日历年预测：{int(has_bbg.sum())} / {len(wind)} 只", flush=True)

    d = wind.copy()
    for col in ("基期提示", "预测来源", "口径", "FY%d来源" % by):
        if col in d:   # Excel 读回的空字符串会变成 NaN，统一还原为空串，避免 NaN != "" 被当成有标记
            d[col] = d[col].fillna("").astype(str)
    cols = [np_ for np_ in [f"NP{by}", f"NP{fy}"] + [f"NP{fy}_{w}前" for w in wins]]
    if "预测来源" not in d or d["预测来源"].eq("").all():
        d["预测来源"] = np.where(d[f"NP{fy}"].notna(), "H股一致预期", "无预测")
    if a.mode == "fill":
        target = d[f"NP{fy}"].isna() & has_bbg & ~d.index.isin(list(DUPLICATE_LINES))
        stem_tag, title_tag = "BBG补缺", "BBG 补缺"
    else:
        target = pd.Series(True, index=d.index) & ~d.index.isin(list(DUPLICATE_LINES))
        stem_tag, title_tag = ("BBG全量", "BBG 全量") if a.mode == "full" else ("BBG相对财年", "BBG 相对财年 FY1→FY2")
    pre = "BBGFY_" if a.mode == "full-fy" else "BBG_"
    if a.mode == "full-fy":
        has_bbg = b[f"BBGFY_NP{fy}"].notna()
    for col in cols:
        d.loc[target, col] = b.loc[target, pre + col] if pre + col in b else np.nan
    if a.mode != "fill":
        d.loc[d.index.isin(list(DUPLICATE_LINES)), cols] = np.nan
    d.loc[target, "预测来源"] = np.where(has_bbg[target], "BBG一致预期", "无预测")
    d.loc[target, "口径"] = b.loc[target, "BBG口径"] if a.mode != "full-fy" else "相对财年"
    # 基期提示按替换后的数（港元）重算被替换的行
    d.loc[target, f"FY{by}原币"] = b.loc[target, f"{pre}NP{by}"]
    d.loc[target, f"FY{fy}原币"] = b.loc[target, f"{pre}NP{fy}"]
    if "基期提示" in d:
        base, nxt = d[f"FY{by}原币"], d[f"FY{fy}原币"]
        flag = (base > 0) & nxt.notna() & (nxt < base * 0.7)
        d.loc[target, "基期提示"] = np.where(flag[target], "基期预期偏高", "")
    print(f"替换 {int((target & has_bbg).sum())} 只（模式 {a.mode}）", flush=True)

    import hk_consensus_2027 as hk
    rows = []
    for name, g in d.groupby("行业", sort=False):
        row = core.summarize(g, name)
        if "基期提示" in g and "折算方式" in g:
            row.update(hk.hk_extra_stats(g))
            extra_flags = [f for f, hit in (("财年错位", row["未日历化利润占比(非12月财年)"] > hk.NON_CAL_SHARE),
                                            ("币种缺失", row["币种未识别市值占比"] > hk.NO_FX_SHARE),
                                            ("基期含一次性", row["基期偏高利润占比"] > hk.ONE_OFF_SHARE)) if hit]
            row["提示"] = "、".join(x for x in [row.get("提示", ""), *extra_flags] if x)
        rows.append(row)

    notes = sheets["统计说明"].iloc[:, 0].dropna().astype(str).tolist()
    extra = [
        (f"{title_tag}：" + ("Wind 无 2027E 预测的公司用 BBG BEST_NET_INCOME 补，个股明细“预测来源”=BBG一致预期；"
                             if a.mode == "fill" else "全部公司改用 BBG BEST_NET_INCOME（BBG 标准口径，部分为经调整）；"
                             + ("“2026E / 2027E”列实为各公司相对财年 FY1 / FY2，未做日历化，用于对照 BBG 指数 FY1/FY2；"
                                if a.mode == "full-fy" else ""))
         + "BBG 金额以 EQY_FUND_CRNCY=HKD 换算港元，市值、行业沿用 Wind 明细。"),
        "BBG 相对财年按 BEST_PERIOD_END_DT 归到截止年份后日历化：CY(Y) = m/12×V(Y) + (1−m/12)×V(Y+1)；"
        "已报财年缺 BBG 值时用 Wind 已报值（折港元）。历史对比日按“前一财年截止 + 3 个月才报出”近似判断是否错开一位。",
        f"同一公司多条股票线只计一条：{'、'.join(f'{k} 以 {v} 为准' for k, v in DUPLICATE_LINES.items())}。",
    ]
    notes = notes[:-1] + extra + notes[-1:]
    stem = Path(a.wind).stem
    cls = "恒生一级" if "恒生一级" in stem else "Wind二级" if "Wind二级" in stem else "行业"
    title = f"港股 · {cls}行业（{title_tag}）"
    out_stem = f"港股_{cls}_{fy}盈利预期_{stem_tag}"
    out = core.write_outputs(rows, [d], notes, out_stem, title)

    if a.mode == "full":   # Wind 与 BBG 逐行业对照
        comp = []
        for name, g in wind.groupby("行业", sort=False):
            gb = d.loc[g.index]
            def growth(x):
                s = x.dropna(subset=[f"NP{by}", f"NP{fy}"])
                return core.positive_denominator_ratio(s[f"NP{fy}"].sum(), s[f"NP{by}"].sum()) - 1
            comp.append({"行业": name, "Wind有预测": int(g[f"NP{fy}"].notna().sum()),
                         "BBG有预测": int(gb[f"NP{fy}"].notna().sum()),
                         f"Wind {fy}E增速": growth(g), f"BBG {fy}E增速": growth(gb),
                         f"Wind {fy}E利润(亿)": g[f"NP{fy}"].sum() / 1e8, f"BBG {fy}E利润(亿)": gb[f"NP{fy}"].sum() / 1e8})
        comp = pd.DataFrame(comp)
        both = wind[[f"NP{fy}"]].join(d[[f"NP{fy}"]], rsuffix="_BBG").dropna()
        both["差异"] = both[f"NP{fy}_BBG"] / both[f"NP{fy}"] - 1
        both = both.join(wind[["SEC_NAME", "行业", "MKT_CAP_ARD"]])
        big = both.reindex(both["差异"].abs().sort_values(ascending=False).index)
        with pd.ExcelWriter(out, engine="openpyxl", mode="a") as xw:
            comp.to_excel(xw, sheet_name="Wind与BBG行业对照", index=False)
            big.rename_axis("股票代码").to_excel(xw, sheet_name="Wind与BBG个股差异")
        print(comp.round(3).to_string(index=False))


if __name__ == "__main__":
    main()
