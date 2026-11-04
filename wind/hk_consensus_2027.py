"""港股分行业一致预期盈利汇总（Wind），口径与申万一级脚本一致，输出 Excel + HTML 看板。

依赖：WindPy、pandas、openpyxl，已登录的 Wind 终端；以及本地财报库（fin_common.py / hk_fx.py 与 SQLite 库）：
  - hk_fx_profile + fx_parity：每家公司的报表币种与人民币中间价（WindPy 只返回报表原币，currencyType 无效，
    2026-10-19 诊断），一致预期按取数日中间价交叉折算为港元；
  - hk_fye：每家公司识别出的财年截止月（公告日 + 历史折算系数，见 hk_calendar.py）。

先预检（核对一致预期的币种与财年标签、行业字段），确认后全量运行：
    python hk_consensus_2027.py --check
    python hk_consensus_2027.py

日历化：一致预期只有整财年数，不能像历史报表那样用季报拼日历期，因此按月份加权：
  CY(Y) = m/12 × FY(Y) + (1 − m/12) × FY(Y+1)，m 为财年截止月；
  Wind 财年标签 N = 截止于 N 年 m 月的财年（与 hk_calendar.py 中“标签月 − 偏移”的归期一致，--check 会用已报财年验证）。
"""
import sqlite3
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
# 财报库代码目录（fin_common.py、hk_fx.py 所在处）；本脚本与它们放在同一目录时可不改
FIN_DIR = Path("/Users/amber/Documents/本地数据库/财报")

# 股票池：恒生综合指数（约 500 只，覆盖港股约 95% 市值）
UNIVERSE_INDEX = "HSCI.HI"

# 行业分类：一次取数，按每种分类各出一套 Excel / 看板；字段取不到的分类跳过并提示
# Wind 行业沿用 wind_fields.json 中 probe 选定的 wicsname2024；恒生行业字段需在 CG 核对（--check 会打印）
CLASSIFICATIONS = {
    "恒生一级": {"field": "industry_hs", "options": "tradeDate={date};industryType=1",
                 "title": "港股 · 恒生一级行业"},
    "Wind二级": {"field": "wicsname2024", "options": "tradeDate={date};industryType=2",
                 "title": "港股 · Wind二级行业"},
}

# 一致预期的币种："report" = 报表原币（需按报表币种折算）；"HKD" = Wind 已折为港元。以 --check 结论为准
CONSENSUS_CCY = "report"
NON_CAL_SHARE = 0.2      # 未能日历化的非 12 月财年利润占比超过此值 -> “财年错位”
NO_FX_SHARE = 0.05       # 币种未识别公司市值占比超过此值 -> “币种缺失”
CHECK_CODES = ["0700.HK", "0939.HK", "2318.HK", "0005.HK", "0992.HK", "9988.HK"]
# EDB 人民币中间价代码 -> 币种（wind_fields.json 注释中列出的部分；其余显示代码）
CCY_NAMES = {"HKD": "港元", "CNY": "人民币", "M0000185": "美元", "M0000186": "欧元",
             "M0000188": "日元(100)", "M0000189": "英镑", "M0000191": "澳元"}
FALLBACK_FX_CODES = ["M0000185", "M0000186", "M0000188", "M0000189", "M0000190", "M0000191", "M0000192",
                     "M0000193", "M0000194", "M0000195", "M0000196", "M0000197", "M0000198"]


# ---------------------------------------------------------------- 本地财报库
def load_fin_modules():
    if str(FIN_DIR) not in sys.path:
        sys.path.insert(0, str(FIN_DIR))
    try:
        import fin_common as C
        import hk_fx
    except ImportError as exc:
        raise RuntimeError(f"找不到财报库模块（{exc}）：请把 FIN_DIR 改为 fin_common.py / hk_fx.py 所在目录") from exc
    return C, hk_fx


def fx_codes(C):
    import json
    cfg = Path(C.__file__).with_name("wind_fields.json")
    try:
        return json.loads(cfg.read_text(encoding="utf-8"))["H"]["fx"]["codes"]
    except (OSError, KeyError, ValueError):
        return FALLBACK_FX_CODES


def ensure_fx_current(w, con, C, hk_fx, log=print):
    """fx_parity 若未覆盖 DATE，则用 EDB 补到今天（与 fin_update_wind.py 同一张表、同样写法）。"""
    hk_fx.ensure_tables(con)
    last = con.execute("SELECT MAX(date) FROM fx_parity WHERE code=?", (hk_fx.HKD_CODE,)).fetchone()[0]
    if last and last >= DATE:
        return

    def edb_series(code, start, end):
        d = w.edb(code, start, end, "Fill=Previous")
        check(d.ErrorCode, f"EDB {code}")
        out = []
        for t, v in zip(d.Times or [], (d.Data or [[]])[0]):
            ds = t.strftime("%Y-%m-%d") if hasattr(t, "strftime") else str(t)[:10]
            out.append((ds, float(v) if v is not None and v == v else None))
        return out

    hk_fx.refresh_fx(con, edb_series, [hk_fx.HKD_CODE] + fx_codes(C), log)


def fx_table(con, hk_fx, codes):
    """每只股票：报表币种、原币→港元系数（取数日中间价）、折算方式。"""
    fx = hk_fx.FX(con)
    prof = {r[0]: r[1:] for r in con.execute(
        "SELECT wind_code, fx_code, direction, last_factor, COALESCE(source,'implied') FROM hk_fx_profile")}
    rows = []
    for c in codes:
        p = prof.get(c)
        if CONSENSUS_CCY == "HKD":
            rows.append((c, "HKD(Wind已折算)", 1.0, "consensus_hkd"))
        elif p and p[0]:
            f = fx.factor(p[0], p[1], 0, DATE)
            name = CCY_NAMES.get(p[0], p[0]) + ("" if p[1] == 1 or p[0] in ("HKD", "CNY") else "(反向报价)")
            rows.append((c, name, f, ("rate_est" if p[3].startswith("est") else "rate") if f else None))
        elif p and p[2]:
            rows.append((c, "未识别(沿用历史系数)", p[2], "approx"))
        else:
            rows.append((c, "未识别", np.nan, None))
    return pd.DataFrame(rows, columns=["code", "报表币种", "折算系数", "折算方式"]).set_index("code")


def fye_months(con, codes):
    """hk_fye：cal=1 用识别出的财年截止月，cal=0 为 12 月；库内没有的公司记为未知。"""
    try:
        got = {r[0]: (r[1], r[2]) for r in con.execute("SELECT wind_code, fye_month, cal FROM hk_fye")}
    except sqlite3.OperationalError:
        got = {}
    return pd.Series([float(got[c][0]) if c in got and got[c][1] else (12.0 if c in got else np.nan)
                      for c in codes], index=codes)


# ---------------------------------------------------------------- 取数与日历化
def wss(w, codes, fields, options, what):
    err, frame = w.wss(codes, fields, options, usedf=True)
    check(err, what)
    return normalize_wss(frame, what).reindex(codes)


def fetch_years(w, codes, years, trade_date, label):
    out = {}
    for year in years:
        e = wss(w, codes, "west_netprofit,west_instnum",
                f"year={year};tradeDate={trade_date};westPeriod=180;unit=1", f"{label} FY{year} 预期")
        out[year] = (pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce"),
                     pd.to_numeric(e["WEST_INSTNUM"], errors="coerce"))
    return out


def calendarize(fy, fy_next, month):
    """返回 (日历年利润, 口径标签)。12 月年结、截止月未知或缺下一财年预测时直接用财年数。"""
    fy, fy_next = fy.astype(float), fy_next.astype(float)
    weight = month / 12
    dec, unknown = month.eq(12), month.isna()
    both = fy.notna() & fy_next.notna()
    value = pd.Series(np.where(dec | unknown | ~both, fy, weight * fy + (1 - weight) * fy_next), index=fy.index)
    basis = pd.Series(np.select([dec, unknown, both], ["12月财年", "财年(截止月未知)", "日历化"],
                                default="财年(缺下一年预测)"), index=fy.index)
    basis[fy.isna()] = ""
    return value, basis


def hk_extra_stats(g):
    f = np_col(FORECAST_YEAR)
    has = g.dropna(subset=[f])
    non_dec = has["财年截止月"].ne(12) & has["财年截止月"].notna()
    not_cal = non_dec & has["口径"].ne("日历化")
    no_fx = g["折算方式"].isna()
    return {
        "非12月财年家数": int(non_dec.sum()),
        "未日历化利润占比(非12月财年)": core.positive_denominator_ratio(has.loc[not_cal, f].sum(),
                                                                       has[f].sum(min_count=1)),
        "币种未识别家数": int(no_fx.sum()),
        "币种未识别市值占比": core.positive_denominator_ratio(g.loc[no_fx, "MKT_CAP_ARD"].sum(),
                                                              g["MKT_CAP_ARD"].sum(min_count=1)),
    }


def hk_notes(revision_dates, universe_size, cls_name):
    cls = CLASSIFICATIONS[cls_name]
    lines = core.note_lines(revision_dates, DATE)
    lines[1] = (f"股票池为{UNIVERSE_INDEX}成分股（{universe_size}只），按{cls_name}分类（字段 {cls['field']}）；"
                "只汇总有 Wind 一致预期覆盖的公司，不代表全行业盈利。")
    lines.insert(2, ("一致预期按报表原币取数（WindPy currencyType 对港股无效），用财报库 hk_fx_profile 识别的报表币种、"
                     f"按 {DATE} 人民币中间价交叉折算为港元；币种未识别且无历史系数的公司不计入。"
                     if CONSENSUS_CCY == "report" else "一致预期经 --check 确认已由 Wind 折为港元，未再折算。")
                 + "金额单位亿港元；市值 mkt_cap_ard 为港元（A+H 公司为各类股份合计市值）。")
    lines.insert(3, "非12月年结公司按财报库 hk_fye 识别的财年截止月 m 日历化：CY(Y) = m/12×FY(Y) + (1−m/12)×FY(Y+1)；"
                    "缺下一财年预测或截止月未知时退回财年口径，个股明细“口径”列逐只标注。"
                    "历史报表的日历对齐用季报/半年报拼接（hk_calendar.py），预测只有整财年数，故用加权近似。")
    lines.insert(-1, f"港股另有提示：未能日历化的非12月财年利润占比>{NON_CAL_SHARE:.0%}为“财年错位”；"
                     f"币种未识别公司市值占比>{NO_FX_SHARE:.0%}为“币种缺失”。")
    return lines


# ---------------------------------------------------------------- 预检
def run_check(w, con, hk_fx):
    print(f"== 股票池 {UNIVERSE_INDEX}")
    err, cons = w.wset("indexconstituent", f"date={DATE};windcode={UNIVERSE_INDEX}", usedf=True)
    print(f"ErrorCode={err}，成分股 {0 if err else len(cons)} 只（恒生综合指数应约 500 只）")
    codes = [] if err else cons["wind_code"].astype(str).str.upper().tolist()

    for name, cfg in CLASSIFICATIONS.items():
        print(f"\n== 行业分类：{name}（{cfg['field']}）")
        e, frame = w.wss(CHECK_CODES, cfg["field"], cfg["options"].format(date=DATE.replace("-", "")), usedf=True)
        print(frame.to_string() if e == 0 else f"ErrorCode={e}：字段不可用，正式运行会跳过该分类")

    prof = pd.read_sql("SELECT wind_code, fx_code FROM hk_fx_profile", con).set_index("wind_code")["fx_code"]
    fye = fye_months(con, codes or list(prof.index))
    last = BASE_YEAR - 1

    def ratio_test(sample, west_year, rpt_year, what):
        if not sample:
            print(f"   {what}：样本为空")
            return None
        est = wss(w, sample, "west_netprofit", f"year={west_year};tradeDate={DATE};westPeriod=180;unit=1", what)
        act = wss(w, sample, "np_belongto_parcomsh", f"unit=1;rptDate={rpt_year}1231;rptType=1", what)
        r = (pd.to_numeric(est.iloc[:, 0], errors="coerce") / pd.to_numeric(act.iloc[:, 0], errors="coerce"))
        r = r[(r > 0) & np.isfinite(r)]
        return r

    print(f"\n== 一致预期币种：FY{last} 一致预期 / FY{last} 已报归母净利润（原币），12 月财年公司")
    fx = hk_fx.FX(con)
    dec = [c for c in (codes or list(prof.index)) if fye.get(c) == 12]
    for code, label in (("CNY", "人民币报表"), ("M0000185", "美元报表"), ("HKD", "港元报表")):
        sample = [c for c in dec if prof.get(c) == code][:30]
        r = ratio_test(sample, last, last, label)
        if r is None or r.empty:
            continue
        g = fx.factor(code, 1, 0, DATE)
        print(f"   {label}：{len(r)} 家，比值中位数 {r.median():.3f}（若 Wind 已折港元，应约 {g:.3f}）")
    print("   -> 人民币/美元报表的比值都接近 1.000：一致预期为报表原币，保持 CONSENSUS_CCY = 'report'；"
          "\n      接近括号里的汇率：Wind 已折港元，改 CONSENSUS_CCY = 'HKD'。")

    print(f"\n== 财年标签：3 月年结公司 FY{BASE_YEAR} 一致预期 vs 已报年报（标签 {BASE_YEAR}-12-31 = 截至 {BASE_YEAR} 年 3 月）")
    mar = [c for c in (codes or list(fye.index)) if fye.get(c) == 3][:20]
    r_same = ratio_test(mar, BASE_YEAR, BASE_YEAR, "同标签")
    r_prev = ratio_test(mar, BASE_YEAR, BASE_YEAR - 1, "错一年")
    if r_same is not None and not r_same.empty and r_prev is not None and not r_prev.empty:
        dev = lambda r: float(np.median(np.abs(np.log(r))))  # noqa: E731
        a, b = dev(r_same), dev(r_prev)
        verdict = ("两者接近，无法区分（样本公司两年利润变化太小），请把结果发给我" if abs(a - b) < 0.005
                   else "约定一致，可直接日历化" if a < b else "标签错了一年，请把结果发给我")
        print(f"   同标签偏差 {a:.3f}，错一年偏差 {b:.3f}（{len(r_same)} 家） -> {verdict}")

    print(f"\n== FY{FORECAST_YEAR + 1} 预期覆盖（日历化需要）")
    nxt = wss(w, CHECK_CODES, "west_netprofit", f"year={FORECAST_YEAR + 1};tradeDate={DATE};westPeriod=180;unit=1", "FY+1")
    print(nxt.to_string())
    no_prof = [c for c in codes if c not in prof.index or pd.isna(prof.get(c))]
    print(f"\n== 财报库覆盖：{len(codes)} 只成分股中，无报表币种识别 {len(no_prof)} 只，"
          f"无财年记录 {int(fye.reindex(codes).isna().sum()) if codes else '—'} 只"
          + (f"（例：{'、'.join(no_prof[:8])}）" if no_prof else ""))
    print("   无币种识别的多为新上市公司：先在财报库跑 fin_update_wind.py hk-estimate / update 补齐。")


# ---------------------------------------------------------------- 主流程
def main(check_only=False):
    from WindPy import w

    C, hk_fx = load_fin_modules()
    core.BASE_YEAR, core.FORECAST_YEAR = BASE_YEAR, FORECAST_YEAR
    core.REVISION_WINDOWS, core.DATE, core.OUT_DIR = REVISION_WINDOWS, DATE, OUT_DIR

    check(w.start().ErrorCode, "Wind 启动")
    if not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")
    con = sqlite3.connect(C.DB_PATH)
    try:
        ensure_fx_current(w, con, C, hk_fx)
        if check_only:
            run_check(w, con, hk_fx)
            return
        run(w, con, hk_fx)
    finally:
        con.close()


def run(w, con, hk_fx):
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

    base = wss(w, codes, "sec_name,mkt_cap_ard", f"tradeDate={DATE};unit=1", "名称/市值")
    d = pd.DataFrame(index=pd.Index(codes))
    d["SEC_NAME"] = base["SEC_NAME"]
    d["MKT_CAP_ARD"] = pd.to_numeric(base["MKT_CAP_ARD"], errors="coerce")
    if d["MKT_CAP_ARD"].isna().all():
        raise RuntimeError(f"{DATE} 市值全部为空，请检查数据日期；未导出本次结果")

    fxs = fx_table(con, hk_fx, codes)
    d = d.join(fxs)
    d["财年截止月"] = fye_months(con, codes)
    missing_fx = d["折算方式"].isna()
    if missing_fx.any():
        print(f"警告：{int(missing_fx.sum())} 只股票无报表币种识别，不计入汇总（市值占比 "
              f"{d.loc[missing_fx, 'MKT_CAP_ARD'].sum() / d['MKT_CAP_ARD'].sum():.1%}）", flush=True)
    if d["财年截止月"].isna().any():
        print(f"警告：{int(d['财年截止月'].isna().sum())} 只股票不在 hk_fye 中，按财年口径处理", flush=True)

    fy_years = (BASE_YEAR, FORECAST_YEAR, FORECAST_YEAR + 1)
    raw = fetch_years(w, codes, fy_years, DATE, "当前")
    factor = d["折算系数"]
    for year in fy_years:
        d[f"FY{year}原币"] = raw[year][0]
    d[f"INST{BASE_YEAR}"] = raw[BASE_YEAR][1]
    d[f"INST{FORECAST_YEAR}"] = raw[FORECAST_YEAR][1]
    hkd = {y: raw[y][0] * factor for y in fy_years}   # 无系数 -> NaN，自然不计入
    d[np_col(BASE_YEAR)], _ = calendarize(hkd[BASE_YEAR], hkd[FORECAST_YEAR], d["财年截止月"])
    d[np_col(FORECAST_YEAR)], d["口径"] = calendarize(hkd[FORECAST_YEAR], hkd[FORECAST_YEAR + 1], d["财年截止月"])
    for window, past_date in revision_dates.items():
        past = fetch_years(w, codes, (FORECAST_YEAR, FORECAST_YEAR + 1), past_date, f"{window}前")
        # 修正只看预期本身的变化：两个日期用同一个（当前）汇率
        d[np_col(FORECAST_YEAR, window)], _ = calendarize(past[FORECAST_YEAR][0] * factor,
                                                          past[FORECAST_YEAR + 1][0] * factor, d["财年截止月"])
    d["请求数据日期"] = DATE
    d["成分股返回日期"] = constituent_dates

    done = 0
    for cls_name, cls in CLASSIFICATIONS.items():
        err, ind = w.wss(codes, cls["field"], cls["options"].format(date=DATE.replace("-", "")), usedf=True)
        if err != 0:
            print(f"跳过 {cls_name}：字段 {cls['field']} 返回 ErrorCode={err}", flush=True)
            continue
        ind = normalize_wss(ind, cls_name).reindex(codes).iloc[:, 0]
        labels = ind.astype("string").str.strip().replace("", pd.NA)
        if labels.isna().all():
            print(f"跳过 {cls_name}：字段 {cls['field']} 全部为空", flush=True)
            continue
        dc = d.copy()
        dc["行业"] = labels.fillna("未分类")
        rows = []
        for name, g in dc.groupby("行业", sort=False):
            row = core.summarize(g, name)
            row.update(hk_extra_stats(g))
            extra = [flag for flag, hit in (("财年错位", row["未日历化利润占比(非12月财年)"] > NON_CAL_SHARE),
                                           ("币种缺失", row["币种未识别市值占比"] > NO_FX_SHARE)) if hit]
            row["提示"] = "、".join(x for x in [row["提示"], *extra] if x)
            rows.append(row)
        print(f"{cls_name}：{len(rows)} 个行业（未分类 {int(labels.isna().sum())} 只）", flush=True)
        core.write_outputs(rows, [dc], hk_notes(revision_dates, len(codes), cls_name),
                           f"港股_{cls_name}_{FORECAST_YEAR}盈利预期", cls["title"])
        done += 1
    if not done:
        raise RuntimeError("所有行业分类字段都不可用，请先运行 --check")


if __name__ == "__main__":
    main(check_only="--check" in sys.argv[1:])
