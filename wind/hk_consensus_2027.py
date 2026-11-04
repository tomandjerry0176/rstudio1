"""港股分行业一致预期盈利汇总（Wind），口径与申万一级脚本一致，输出 Excel + HTML 看板。

运行环境需要 WindPy、pandas、openpyxl，并已登录可用的 Wind 终端。
先运行预检，确认行业、币种、财年字段在你的终端上返回正确：
    python hk_consensus_2027.py --check
确认无误后全量运行：
    python hk_consensus_2027.py

港股相对 A 股多出的三处口径处理：
1. 币种：利润与市值统一按 CURRENCY 取数（Wind currencyType 参数），避免人民币/美元报表直接相加。
2. 财年：非 12 月年结公司按财年截止月日历化，CY = m/12 × FY(Y) + (1 − m/12) × FY(Y+1)。
3. 行业：用字段给每只股票分类（恒生一级或 Wind 二级），而不是用行业指数成分股。
"""
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import sw1_consensus_2027 as core
from sw1_consensus_2027 import check, normalize_wss, np_col

DATE = "2026-11-03"
BASE_YEAR = 2026
FORECAST_YEAR = 2027
REVISION_WINDOWS = {"4周": 20, "13周": 65}
OUT_DIR = Path("/Users/amber/Documents/本地数据库/财报")

# 股票池：恒生综合指数，覆盖港股约 95% 市值；如需全部港股可换成板块成分股
UNIVERSE_INDEX = "HSCI.HI"

# 行业分类字段（字段名与 industryType 需在 Wind 代码生成器 CG 中核对，--check 会打印返回值）
CLASSIFICATION = "恒生一级"
CLASSIFICATIONS = {
    "恒生一级": {"field": "industry_hs", "options": "industryType=1", "title": "港股 · 恒生一级行业"},
    "Wind二级": {"field": "industry_gics", "options": "industryType=2", "title": "港股 · Wind二级行业"},
}

CURRENCY = "HKD"
CURRENCY_OPT = f"currencyType={CURRENCY}"
FYE_FIELD = "fiscalyearend"          # 财年截止日，需在 CG 中核对
CALENDARIZE = True
NON_CAL_SHARE = 0.2                  # 未能日历化的非 12 月财年利润占比超过此值 -> 提示“财年错位”

# 预检样本：腾讯(人民币报表)、建行(A+H)、平安(A+H)、汇丰(美元报表)、联想(3 月年结)
CHECK_CODES = ["0700.HK", "0939.HK", "2318.HK", "0005.HK", "0992.HK"]


def wss(w, codes, fields, options, what):
    err, frame = w.wss(codes, fields, options, usedf=True)
    check(err, what)
    return normalize_wss(frame, what).reindex(codes)


def fye_month(value):
    """把 Wind 返回的财年截止日（日期、'12-31'、'1231'、'2026-03-31' 等）解析为月份。"""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return np.nan
    if hasattr(value, "month"):
        return float(value.month)
    tokens = re.findall(r"\d+", str(value))
    month = np.nan
    if len(tokens) >= 2:
        month = int(tokens[1]) if len(tokens[0]) == 4 else int(tokens[0])
    elif len(tokens) == 1:
        t = tokens[0]
        if len(t) in (3, 4):
            month = int(t[:-2])
        elif len(t) <= 2:
            month = int(t)
    return float(month) if 1 <= month <= 12 else np.nan


def calendarize(fy, fy_next, month):
    """返回 (日历年利润, 口径标签)。12 月年结或截止月未知时直接用财年数。"""
    fy, fy_next = fy.astype(float), fy_next.astype(float)
    if not CALENDARIZE:
        return fy, pd.Series("财年", index=fy.index)
    weight = month / 12
    dec = month.eq(12)
    unknown = month.isna()
    both = fy.notna() & fy_next.notna()
    value = pd.Series(np.where(dec | unknown | ~both, fy, weight * fy + (1 - weight) * fy_next),
                      index=fy.index)
    basis = pd.Series(np.select(
        [dec, unknown, both],
        ["12月财年", "财年(截止月未知)", "日历化"],
        default="财年(缺下一年预测)",
    ), index=fy.index)
    basis[fy.isna()] = ""
    return value, basis


def fetch_years(w, codes, years, trade_date, label):
    """按财年取一致预期净利润（统一币种）与机构数。"""
    out = {}
    for year in years:
        e = wss(w, codes, "west_netprofit,west_instnum",
                f"year={year};tradeDate={trade_date};westPeriod=180;unit=1;{CURRENCY_OPT}",
                f"{label} FY{year} 预期")
        out[year] = (pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce"),
                     pd.to_numeric(e["WEST_INSTNUM"], errors="coerce"))
    return out


def hk_extra_stats(g):
    """港股特有：非 12 月财年家数、未能日历化的利润占比。"""
    f = np_col(FORECAST_YEAR)
    has = g.dropna(subset=[f])
    non_dec = has["财年截止月"].ne(12) & has["财年截止月"].notna()
    not_cal = non_dec & has["口径"].ne("日历化")
    total = has[f].sum(min_count=1)
    return {
        "非12月财年家数": int(non_dec.sum()),
        "未日历化利润占比(非12月财年)": core.positive_denominator_ratio(
            has.loc[not_cal, f].sum(), total
        ),
    }


def hk_notes(revision_dates, universe_size):
    cls = CLASSIFICATIONS[CLASSIFICATION]
    lines = core.note_lines(revision_dates, DATE)
    lines[1] = (f"股票池为{UNIVERSE_INDEX}成分股（{universe_size}只），按{CLASSIFICATION}分类"
                f"（字段 {cls['field']}，{cls['options']}）；只汇总有 Wind 一致预期覆盖的公司，不代表全行业盈利。")
    lines.insert(2, f"利润与市值按 {CURRENCY_OPT} 统一为{CURRENCY}，金额单位亿{CURRENCY}；"
                    "A+H 公司市值使用 mkt_cap_ard（各类股份按各自股价计算的总市值），与全公司归母净利润匹配。")
    lines.insert(3, ("非12月年结公司已日历化：CY(Y) = m/12×FY(Y) + (1−m/12)×FY(Y+1)，m 为财年截止月；"
                     "缺少下一财年预测时退回财年口径，个股明细“口径”列逐只标注。")
                 if CALENDARIZE else "未做日历化，非12月年结公司按财年口径直接汇总。")
    lines.insert(-1, f"港股提示标签另含“财年错位”：未能日历化的非12月财年利润占行业比重>{NON_CAL_SHARE:.0%}。")
    return lines


def run_check(w):
    cls_fields = {k: v for k, v in CLASSIFICATIONS.items()}
    print(f"== 股票池 {UNIVERSE_INDEX}")
    err, cons = w.wset("indexconstituent", f"date={DATE};windcode={UNIVERSE_INDEX}", usedf=True)
    print(f"ErrorCode={err}，成分股 {0 if err else len(cons)} 只（恒生综合指数应约 500 只）")

    def show(title, fields, options):
        print(f"\n== {title}\n   fields={fields}  options={options}")
        err, frame = w.wss(CHECK_CODES, fields, options, usedf=True)
        if err != 0:
            print(f"   ErrorCode={err}：字段或参数可能不对，请在 CG 中核对")
            return None
        print(frame.to_string())
        return frame

    show("名称 / 市值（统一币种）/ 财年截止日", f"sec_name,mkt_cap_ard,{FYE_FIELD}",
         f"tradeDate={DATE};unit=1;{CURRENCY_OPT}")
    show("市值（不带币种参数，对比用）", "mkt_cap_ard", f"tradeDate={DATE};unit=1")
    for name, cfg in cls_fields.items():
        show(f"行业分类：{name}", cfg["field"], f"tradeDate={DATE};{cfg['options']}")
    a = show(f"FY{FORECAST_YEAR} 预期净利润（统一为{CURRENCY}）", "west_netprofit",
             f"year={FORECAST_YEAR};tradeDate={DATE};westPeriod=180;unit=1;{CURRENCY_OPT}")
    b = show(f"FY{FORECAST_YEAR} 预期净利润（原始币种）", "west_netprofit",
             f"year={FORECAST_YEAR};tradeDate={DATE};westPeriod=180;unit=1")
    if a is not None and b is not None:
        ratio = a.iloc[:, 0] / b.iloc[:, 0]
        print("\n== 统一币种 / 原始币种 比值")
        print(ratio.round(3).to_string())
    show(f"FY{FORECAST_YEAR + 1} 预期覆盖（日历化需要）", "west_netprofit,west_instnum",
         f"year={FORECAST_YEAR + 1};tradeDate={DATE};westPeriod=180;unit=1;{CURRENCY_OPT}")
    print(f"""
== 怎么判断
1. 行业分类：应返回中文行业名（如 资讯科技业 / 金融业）。为空或报错 -> 在 CG 里查“恒生行业”字段名，改 CLASSIFICATIONS。
2. 币种：比值列 0700/0939/2318 应约等于人民币兑港币汇率（约 1.08–1.10），0005 应约 7.8（美元报表）。
   若全部是 1.000，说明 currencyType 对一致预期字段无效，需要改成手动换汇，把结果发给我。
3. 财年：0992（联想）财年截止日应是 3 月，其余为 12 月。为空 -> 在 CG 查“财年截止日”字段，改 FYE_FIELD。
4. 市值：2318/0939 带币种的市值应是 A+H 合计（明显大于只算 H 股的市值）。
5. FY{FORECAST_YEAR + 1}：大市值公司应大多有值，否则 3 月年结公司无法日历化，会退回财年口径。""")


def main(check_only=False):
    from WindPy import w

    # 让共享模块使用港股的年份 / 窗口 / 日期 / 输出目录
    core.BASE_YEAR, core.FORECAST_YEAR = BASE_YEAR, FORECAST_YEAR
    core.REVISION_WINDOWS, core.DATE, core.OUT_DIR = REVISION_WINDOWS, DATE, OUT_DIR

    check(w.start().ErrorCode, "Wind 启动")
    if not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")
    if check_only:
        run_check(w)
        return
    cls = CLASSIFICATIONS[CLASSIFICATION]

    revision_dates = {win: core.offset_trade_date(w, days) for win, days in REVISION_WINDOWS.items()}
    print(f"预期修正对比日期：{revision_dates}", flush=True)

    err, cons = w.wset("indexconstituent", f"date={DATE};windcode={UNIVERSE_INDEX}", usedf=True)
    check(err, f"{UNIVERSE_INDEX} 成分股")
    cons = cons.copy()
    cons.columns = cons.columns.astype(str).str.lower()
    core.require_columns(cons, ["wind_code"], f"{UNIVERSE_INDEX} 成分股")
    codes = cons["wind_code"].dropna().astype(str).str.strip().str.upper()
    codes = codes[codes.ne("")].drop_duplicates().tolist()
    if not codes:
        raise RuntimeError(f"{UNIVERSE_INDEX} 成分股为空，停止导出")
    constituent_dates = "未返回日期"
    if "date" in cons.columns:
        dates = pd.to_datetime(cons["date"], errors="coerce").dropna()
        constituent_dates = ",".join(sorted(set(dates.dt.strftime("%Y-%m-%d"))))
    print(f"{UNIVERSE_INDEX} 成分股 {len(codes)} 只", flush=True)

    base = wss(w, codes, f"sec_name,mkt_cap_ard,{FYE_FIELD}",
               f"tradeDate={DATE};unit=1;{CURRENCY_OPT}", "名称/市值/财年")
    core.require_columns(base, ["SEC_NAME", "MKT_CAP_ARD", FYE_FIELD.upper()], "名称/市值/财年")
    ind = wss(w, codes, cls["field"], f"tradeDate={DATE};{cls['options']}", "行业分类")

    d = pd.DataFrame(index=pd.Index(codes))
    d["SEC_NAME"] = base["SEC_NAME"]
    d["MKT_CAP_ARD"] = pd.to_numeric(base["MKT_CAP_ARD"], errors="coerce")
    if d["MKT_CAP_ARD"].isna().all():
        raise RuntimeError(f"{DATE} 市值全部为空，请检查数据日期；未导出本次结果")
    d["财年截止月"] = base[FYE_FIELD.upper()].map(fye_month)
    d["行业"] = ind.iloc[:, 0].astype("string").str.strip().replace("", pd.NA).fillna("未分类")
    unclassified = int(d["行业"].eq("未分类").sum())
    if unclassified == len(d):
        raise RuntimeError(f"行业字段 {cls['field']} 全部为空，请先运行 --check 核对字段")
    if unclassified:
        print(f"警告：{unclassified} 只股票无{CLASSIFICATION}分类，归入“未分类”", flush=True)
    if d["财年截止月"].isna().mean() > 0.5:
        print(f"警告：过半股票未取到财年截止月（字段 {FYE_FIELD}），日历化基本不会生效", flush=True)

    fy_years = (BASE_YEAR, FORECAST_YEAR, FORECAST_YEAR + 1)
    raw = fetch_years(w, codes, fy_years, DATE, "当前")
    for year in fy_years:
        d[f"FY{year}原始"] = raw[year][0]
    d[f"INST{BASE_YEAR}"] = raw[BASE_YEAR][1]
    d[f"INST{FORECAST_YEAR}"] = raw[FORECAST_YEAR][1]
    d[np_col(BASE_YEAR)], _ = calendarize(raw[BASE_YEAR][0], raw[FORECAST_YEAR][0], d["财年截止月"])
    d[np_col(FORECAST_YEAR)], d["口径"] = calendarize(
        raw[FORECAST_YEAR][0], raw[FORECAST_YEAR + 1][0], d["财年截止月"]
    )
    for window, past_date in revision_dates.items():
        past = fetch_years(w, codes, (FORECAST_YEAR, FORECAST_YEAR + 1), past_date, f"{window}前")
        d[np_col(FORECAST_YEAR, window)], _ = calendarize(
            past[FORECAST_YEAR][0], past[FORECAST_YEAR + 1][0], d["财年截止月"]
        )
    d["请求数据日期"] = DATE
    d["成分股返回日期"] = constituent_dates

    rows = []
    for name, g in d.groupby("行业", sort=False):
        row = core.summarize(g, name)
        row.update(hk_extra_stats(g))
        if row["未日历化利润占比(非12月财年)"] > NON_CAL_SHARE:
            row["提示"] = "、".join(x for x in (row["提示"], "财年错位") if x)
        rows.append(row)
        print(f"完成 {name}: {len(g)} 只", flush=True)

    core.write_outputs(rows, [d], hk_notes(revision_dates, len(codes)),
                       f"港股_{CLASSIFICATION}_{FORECAST_YEAR}盈利预期", cls["title"])


if __name__ == "__main__":
    main(check_only="--check" in sys.argv[1:])
