"""港股分行业一致预期盈利汇总（Wind），口径与申万一级脚本一致，输出 Excel + HTML 看板。

依赖：WindPy、pandas、openpyxl，已登录的 Wind 终端；以及本地财报库（fin_common.py / hk_fx.py 与 SQLite 库）：
  - hk_fx_profile + fx_parity：每家公司的报表币种与人民币中间价（WindPy 只返回报表原币，currencyType 无效，
    2026-10-19 诊断），一致预期按取数日中间价交叉折算为港元；
  - hk_fye：每家公司识别出的财年截止月（公告日 + 历史折算系数，见 hk_calendar.py）。

先预检（核对一致预期的币种与财年标签、行业字段），确认后全量运行：
    python hk_consensus_2027.py --check
    python hk_consensus_2027.py
默认只读打开财报库，用库内 fx_parity 不晚于 DATE 的最新中间价；加 --refresh-fx 才用 EDB 补汇率（写 fx_parity）。
共享模块 sw1_consensus_2027.py 为新版（含 write_outputs）时输出交互看板与提示标签，旧版时退回静态 HTML 表格。

日历化：一致预期只有整财年数，不能像历史报表那样用季报拼日历期，因此按月份加权：
  CY(Y) = m/12 × FY(Y) + (1 − m/12) × FY(Y+1)，m 为财年截止月；
  Wind 财年标签 N = 截止于 N 年 m 月的财年（与 hk_calendar.py 中“标签月 − 偏移”的归期一致，--check 会用已报财年验证）。
"""
import argparse
import html
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# 申万汇总脚本保留在财报输出目录；财报库模块位于本脚本目录。
CORE_DIR = Path("/Users/amber/Documents/本地数据库/财报")
if str(CORE_DIR) not in sys.path:
    sys.path.append(str(CORE_DIR))
import sw1_consensus_2027 as core
from sw1_consensus_2027 import check, normalize_wss, np_col

DATE = "2026-11-03"
BASE_YEAR = 2026
FORECAST_YEAR = 2027
REVISION_WINDOWS = {"4周": 20, "13周": 65}
OUT_DIR = Path("/Users/amber/Documents/本地数据库/财报")
# 财报库代码目录（fin_common.py、hk_fx.py 所在处）；本脚本与它们放在同一目录时可不改
FIN_DIR = Path(__file__).resolve().parent

# 股票池：恒生综合指数（约 500 只，覆盖港股约 95% 市值）
UNIVERSE_INDEX = "HSCI.HI"

# 行业分类：一次取数，按每种分类各出一套 Excel / 看板；字段取不到的分类跳过并提示
# Wind 行业沿用 wind_fields.json 中 probe 选定的 wicsname2024；恒生行业级别参数为 category。
CLASSIFICATIONS = {
    "恒生一级": {"field": "industry_hs", "options": "tradeDate={date};category=1",
                 "title": "港股 · 恒生一级行业"},
    "Wind二级": {"field": "wicsname2024", "options": "tradeDate={date};industryType=2",
                 "title": "港股 · Wind二级行业"},
}

# 一致预期的币种："report" = 报表原币（需按报表币种折算）；"HKD" = Wind 已折为港元。以 --check 结论为准
CONSENSUS_CCY = "report"
NON_CAL_SHARE = 0.2      # 未能日历化的非 12 月财年利润占比超过此值 -> “财年错位”
NO_FX_SHARE = 0.05       # 币种未识别公司市值占比超过此值 -> “币种缺失”
STALE_FX_DAYS = 7        # 只读模式下库内最新汇率早于 DATE 超过此天数时警告

# 手动指定报表币种：财报库尚无币种识别的新上市公司（优先级高于 hk_fx_profile）。值为 (EDB 中间价代码, 方向)，
# 代码可用 "HKD"、"CNY" 或 wind_fields.json 中的 M00001xx；方向 1 = 中间价为“1 外币 = x 人民币”。
# 财报库补齐后（fin_update_wind.py hk-estimate / update）可删掉对应条目。
MANUAL_FX = {
    "0625.HK": ("M0000185", 1),   # 希音：美元报表（2026-11-05 用户指定）
}
# 手动指定财年截止月（hk_fye 无记录的公司）；不填则按财年口径直接使用
MANUAL_FYE = {}

# 基期一次性因素：个股基期（FY{BASE}）利润明显高于下一财年时标注
ONE_OFF_DROP = 0.30           # 基期为一致预期：下一年较基期下降超过 30%
ONE_OFF_DROP_REPORTED = 0.05  # 基期为已报利润：下一年预期低于已报超过 5%（已报含一次性损益，预期通常不含）
ONE_OFF_SHARE = 0.10          # 被标注公司基期利润占行业基期利润超过此值 -> 行业提示“基期含一次性”
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


def fx_rate_date(con, hk_fx):
    """实际使用的中间价日期：fx_parity 中不晚于 DATE 的最新日期。"""
    return con.execute("SELECT MAX(date) FROM fx_parity WHERE code=? AND date<=?", (hk_fx.HKD_CODE, DATE)).fetchone()[0]


def ensure_fx_current(w, con, C, hk_fx, refresh=False, log=print):
    """返回实际使用的中间价日期。refresh=True 且库内未覆盖 DATE 时用 EDB 补齐（写 fx_parity，与 fin_update_wind.py 同表同法）；
    默认只读，不写库，汇率偏旧时只警告。"""
    if refresh:
        hk_fx.ensure_tables(con)
    last = con.execute("SELECT MAX(date) FROM fx_parity WHERE code=?", (hk_fx.HKD_CODE,)).fetchone()[0]
    if not refresh or (last and last >= DATE):
        used = fx_rate_date(con, hk_fx)
        if used is None:
            raise RuntimeError(f"fx_parity 中没有 {DATE} 及以前的港币中间价：请加 --refresh-fx 或先跑 fin_update_wind.py update")
        lag = (pd.Timestamp(DATE) - pd.Timestamp(used)).days
        if lag > STALE_FX_DAYS:
            log(f"警告：库内最新中间价为 {used}，比 {DATE} 早 {lag} 天；需要时加 --refresh-fx")
        return used

    def edb_series(code, start, end):
        d = w.edb(code, start, end, "Fill=Previous")
        check(d.ErrorCode, f"EDB {code}")
        out = []
        for t, v in zip(d.Times or [], (d.Data or [[]])[0]):
            ds = t.strftime("%Y-%m-%d") if hasattr(t, "strftime") else str(t)[:10]
            out.append((ds, float(v) if v is not None and v == v else None))
        return out

    hk_fx.refresh_fx(con, edb_series, [hk_fx.HKD_CODE] + fx_codes(C), log)
    return fx_rate_date(con, hk_fx)


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
        elif c in MANUAL_FX:
            code, direction = MANUAL_FX[c]
            f = fx.factor(code, direction, 0, DATE)
            rows.append((c, CCY_NAMES.get(code, code) + "(手动指定)", f, "manual" if f else None))
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
    out = pd.Series([float(got[c][0]) if c in got and got[c][1] else (12.0 if c in got else np.nan)
                     for c in codes], index=codes)
    for c, m in MANUAL_FYE.items():
        if c in out.index:
            out[c] = float(m)
    return out


# ---------------------------------------------------------------- 取数与日历化
def wss(w, codes, fields, options, what):
    err, frame = w.wss(codes, fields, options, usedf=True)
    check(err, what)
    frame = normalize_wss(frame, what).reindex(codes)
    core.require_columns(frame, fields.upper().split(","), what)
    return frame


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


def base_one_off(d):
    """个股基期提示：基期财年利润（原币）明显高于下一财年，可能含一次性收益。"""
    base, nxt = d[f"FY{BASE_YEAR}原币"], d[f"FY{FORECAST_YEAR}原币"]
    reported = d[f"FY{BASE_YEAR}来源"].eq("已报")
    keep = np.where(reported, 1 - ONE_OFF_DROP_REPORTED, 1 - ONE_OFF_DROP)
    hit = (base > 0) & nxt.notna() & (nxt < base * keep)
    return pd.Series(np.select([hit & reported, hit], ["已报基期高于下年预期", "基期预期偏高"], default=""),
                     index=d.index)


def hk_extra_stats(g):
    f = np_col(FORECAST_YEAR)
    b = np_col(BASE_YEAR)
    both = g.dropna(subset=[b, f])
    flagged = both["基期提示"].ne("")
    rest = both[~flagged]
    top = both[flagged].sort_values(b, ascending=False)
    one_off = {
        "基期偏高公司": "、".join(top["SEC_NAME"].fillna(pd.Series(top.index, index=top.index)).astype(str)),
        "基期偏高利润占比": core.positive_denominator_ratio(both.loc[flagged, b].sum(), both[b].sum(min_count=1)),
        "剔除基期偏高公司后增速": (core.positive_denominator_ratio(rest[f].sum(min_count=1), rest[b].sum(min_count=1)) - 1
                                  if flagged.any() else float("nan")),
    }
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
        **one_off,
    }


def hk_notes(revision_dates, universe_size, cls_name, rate_date):
    cls = CLASSIFICATIONS[cls_name]
    lines = [
        f"请求数据日期：{DATE}；成分股实际返回日期见个股明细。",
        "",  # 下方填入港股股票池与分类口径
        "两年利润与增速使用共同样本；基期合计利润不为正时增速留空。",
        "PE使用市值与预测利润均非空的同一批股票，保留亏损公司；预测合计利润不为正时PE留空。",
        "覆盖率分母仅含已取得的市值；存在市值缺失时不是完整行业市值覆盖率。",
        "机构数使用west_instnum；year参数是否适用于该字段仍需核对终端定义。",
        "预期修正对比日期：" + "；".join(f"{win}前={dt}" for win, dt in revision_dates.items()) + "。",
        "修正使用当前成分股、两个日期都有预测的共同样本和相同汇率；对比日合计利润不为正时留空。",
        "上调净占比=(上调-下调)/修正样本家数；相对变动小于百万分之一视为不变。",
        f"前{core.TOP_N}公司占比以全部有预测公司的利润合计为分母，含亏损公司时可能超过100%；分母不为正时留空。",
        "个股增速中位数仅统计基期预测利润大于零的公司。",
        "空白表示缺失或不适用，不以零填补。",
    ]
    lines[1] = (f"股票池为{UNIVERSE_INDEX}成分股（{universe_size}只），按{cls_name}分类（字段 {cls['field']}）；"
                "只汇总有 Wind 一致预期覆盖的公司，不代表全行业盈利。")
    lines.insert(2, ("一致预期按报表原币取数（WindPy currencyType 对港股无效），用财报库 hk_fx_profile 识别的报表币种、"
                     f"按人民币中间价（{rate_date}）交叉折算为港元；币种未识别的公司不计入，"
                     "无当前币种但有历史系数的沿用历史系数（个股明细“折算方式”为 approx）。"
                     if CONSENSUS_CCY == "report" else "一致预期经 --check 确认已由 Wind 折为港元，未再折算。")
                 + "金额单位亿港元；市值 mkt_cap_ard 为港元（A+H 公司为各类股份合计市值）。")
    lines.insert(3, f"基期财年已结束、Wind 不再保留一致预期的公司（如 3 月年结的 FY{BASE_YEAR}），用已报归母净利润（原币）"
                    f"补足基期，个股明细“FY{BASE_YEAR}来源”列标注。"
                    "非12月年结公司按财报库 hk_fye 识别的财年截止月 m 日历化：CY(Y) = m/12×FY(Y) + (1−m/12)×FY(Y+1)；"
                    "缺下一财年预测或截止月未知时退回财年口径，个股明细“口径”列逐只标注。"
                    "历史报表的日历对齐用季报/半年报拼接（hk_calendar.py），预测只有整财年数，故用加权近似。")
    if hasattr(core, "flag_text"):
        lines.insert(-1, "增量 = 共同样本两年预测利润之差；扭亏贡献占比 = 基期亏损公司贡献的增量 / 行业总增量；"
                         "最大增量公司为共同样本中增量最大的一家，剔除后增速为去掉该公司后的共同样本增速。")
        lines.insert(-1, f"提示标签：有预测家数<{core.SMALL_SAMPLE}为小样本；基期合计亏损为基期亏损；"
                         f"扭亏贡献>{core.TURNAROUND_SHARE:.0%}为扭亏驱动；单一公司增量占比>{core.SINGLE_NAME_SHARE:.0%}"
                         f"为单一公司主导；前{core.TOP_N}公司利润占比>{core.CONCENTRATION:.0%}为集中度高；"
                         f"未能日历化的非12月财年利润占比>{NON_CAL_SHARE:.0%}为财年错位；"
                         f"币种未识别公司市值占比>{NO_FX_SHARE:.0%}为币种缺失；"
                         f"基期偏高公司的基期利润占行业>{ONE_OFF_SHARE:.0%}为基期含一次性。")
    else:
        lines.insert(-1, f"港股提示标签：未能日历化的非12月财年利润占比>{NON_CAL_SHARE:.0%}为“财年错位”；"
                         f"币种未识别公司市值占比>{NO_FX_SHARE:.0%}为“币种缺失”；"
                         f"基期偏高公司的基期利润占行业>{ONE_OFF_SHARE:.0%}为“基期含一次性”。")
    lines.insert(-1, f"基期偏高（个股明细“基期提示”）：基期为一致预期且 FY{FORECAST_YEAR} 预期较 FY{BASE_YEAR} 下降超过"
                     f"{ONE_OFF_DROP:.0%}，或基期为已报利润且 FY{FORECAST_YEAR} 预期低于已报超过{ONE_OFF_DROP_REPORTED:.0%}"
                     "（已报含一次性损益、预期通常不含），可能含出售资产、投资收益等一次性因素；按原币财年数判断，只作提示、不剔除。"
                     "行业汇总给出基期偏高公司、其基期利润占比和剔除后增速。")
    if MANUAL_FX or MANUAL_FYE:
        lines.insert(-1, "手动指定："
                         + "；".join([f"{c} 报表币种 {CCY_NAMES.get(v[0], v[0])}" for c, v in MANUAL_FX.items()]
                                    + [f"{c} 财年截止月 {m}" for c, m in MANUAL_FYE.items()])
                         + "（财报库尚无识别，个股明细“折算方式”为 manual）。")
    return lines


def write_outputs(rows, details, notes, stem, title):
    """优先用共享模块的 write_outputs（Excel + 交互看板）；共享模块为旧版时退回本地静态导出。"""
    if hasattr(core, "write_outputs"):
        return core.write_outputs(rows, details, notes, stem, title)
    print("提示：sw1_consensus_2027.py 为旧版（无 write_outputs），输出静态 HTML；"
          "更新 sw1_consensus_2027.py / build_dashboard.py / dashboard_template.html 后可得交互看板", flush=True)
    return write_static(rows, details, notes, stem, title)


def write_static(rows, details, notes, stem, title):
    summary = pd.DataFrame(rows).sort_values(
        f"{FORECAST_YEAR}E增速(共同样本)", ascending=False, na_position="last")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{stem}_{DATE}"
    with pd.ExcelWriter(path.with_suffix(".xlsx"), engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="行业汇总", index=False)
        pd.concat(details).rename_axis("股票代码").to_excel(writer, sheet_name="个股明细")
        pd.DataFrame({"统计说明": notes}).to_excel(writer, sheet_name="统计说明", index=False)
        ws = writer.sheets["行业汇总"]
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        for column, name in enumerate(summary.columns, 1):
            if any(key in name for key in ("覆盖率", "增速", "修正_", "占比")):
                for cells in ws.iter_rows(min_row=2, min_col=column, max_col=column):
                    cells[0].number_format = "0.00%"
    shown = summary.copy()
    for name in shown.columns:
        if pd.api.types.is_numeric_dtype(shown[name]):
            percent = any(key in name for key in ("覆盖率", "增速", "修正_", "占比"))
            shown[name] = shown[name].map(
                lambda value: "" if pd.isna(value) else (f"{value:.2%}" if percent else f"{value:,.2f}"))
    page = ("<!doctype html><html lang='zh-CN'><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>{html.escape(title)}</title><style>"
            "body{font-family:system-ui,sans-serif;margin:24px;color:#172b4d}"
            "table{border-collapse:collapse;white-space:nowrap}td,th{padding:8px;border:1px solid #ddd}"
            "th{background:#eef3f8}li{margin:8px 0}</style>"
            f"<h1>{html.escape(title)} · {FORECAST_YEAR}盈利预期</h1><p>数据日期：{DATE}；金额单位：亿港元</p>"
            "<div style='overflow:auto'>" + shown.to_html(index=False, escape=True) + "</div><h2>统计说明</h2><ul>"
            + "".join(f"<li>{html.escape(note)}</li>" for note in notes) + "</ul></html>")
    path.with_suffix(".html").write_text(page, encoding="utf-8")
    print(f"已保存：{path.with_suffix('.xlsx')} / {path.with_suffix('.html')}", flush=True)


# ---------------------------------------------------------------- 预检
def run_check(w, con, hk_fx):
    print(f"== 股票池 {UNIVERSE_INDEX}")
    err, cons = w.wset("indexconstituent", f"date={DATE};windcode={UNIVERSE_INDEX}", usedf=True)
    check(err, f"{UNIVERSE_INDEX} 成分股预检")
    cons.columns = cons.columns.astype(str).str.lower()
    core.require_columns(cons, ["wind_code"], "成分股预检")
    codes = cons["wind_code"].dropna().astype(str).str.strip().str.upper().drop_duplicates().tolist()
    if not codes:
        raise RuntimeError("成分股预检为空")
    print(f"ErrorCode={err}，成分股 {len(codes)} 只")

    for name, cfg in CLASSIFICATIONS.items():
        print(f"\n== 行业分类：{name}（{cfg['field']}）")
        e, frame = w.wss(CHECK_CODES, cfg["field"], cfg["options"].format(date=DATE.replace("-", "")), usedf=True)
        print(frame.to_string() if e == 0 else f"ErrorCode={e}：字段不可用，正式运行会跳过该分类")
        if e == 0 and frame.isna().all().all():
            print("警告：字段全部为空，未通过分类预检")

    prof = pd.read_sql("SELECT wind_code, fx_code FROM hk_fx_profile", con).set_index("wind_code")["fx_code"]
    fye = fye_months(con, codes)
    fx = hk_fx.FX(con)
    prev = BASE_YEAR - 1

    def west(sample, year):
        e = wss(w, sample, "west_netprofit", f"year={year};tradeDate={DATE};westPeriod=180;unit=1", f"FY{year} 预期")
        return pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce")

    def actual(sample, label_year):
        a = wss(w, sample, "np_belongto_parcomsh", f"unit=1;rptDate={label_year}1231;rptType=1", f"{label_year} 年报")
        return pd.to_numeric(a["NP_BELONGTO_PARCOMSH"], errors="coerce")

    # 已报年份的一致预期 Wind 不再保留（上一轮 FY{prev} 预期/已报配对全空即此原因），改用下一年预期对比已报
    print(f"\n== 一致预期币种：FY{BASE_YEAR} 一致预期 / FY{prev} 已报归母净利润（原币），12 月财年、两年均为正")
    print("   原币口径下比值 ≈ 1+增速；若 Wind 已折港元，比值 ≈ 交叉汇率×(1+增速)。美元报表最好分辨（汇率约 7.8）。")
    dec = [c for c in codes if fye.get(c) == 12]
    medians = {}
    for code, label in (("M0000185", "美元报表"), ("CNY", "人民币报表"), ("HKD", "港元报表")):
        sample = [c for c in dec if prof.get(c) == code][:40]
        if not sample:
            print(f"   {label}：样本为空")
            continue
        e, a = west(sample, BASE_YEAR), actual(sample, prev)
        ok = (e > 0) & (a > 0)
        r = (e / a)[ok]
        g = fx.factor(code, 1, 0, DATE)
        line = f"   {label}：取样 {len(sample)}，预期非空 {int(e.notna().sum())}，已报非空 {int(a.notna().sum())}，有效 {len(r)}"
        if len(r):
            line += f"，比值中位数 {r.median():.3f}（交叉汇率 {g:.3f}）" if g else f"，比值中位数 {r.median():.3f}（汇率缺失）"
        print(line)
        if len(r) >= 5 and g:
            medians[code] = (float(r.median()), g)
    verdict = None
    if "M0000185" in medians:
        m, g = medians["M0000185"]
        verdict = "report" if m < g / 2.5 else "HKD" if m > g / 2 else None
    elif "CNY" in medians and "HKD" in medians:
        rel = medians["CNY"][0] / medians["HKD"][0]
        print(f"   人民币组 / 港元组比值 {rel:.3f}（原币口径约 1，港元口径约 {medians['CNY'][1]:.3f}，受两组增速差影响，仅供参考）")
    if verdict:
        same = verdict == CONSENSUS_CCY
        print(f"   -> 结论：一致预期为{'报表原币' if verdict == 'report' else '港元'}；"
              + ("与当前 CONSENSUS_CCY 一致。" if same else f"与当前 CONSENSUS_CCY='{CONSENSUS_CCY}' 不一致，应改为 '{verdict}'。"))
    else:
        print("   警告：美元报表有效样本不足，币种口径未能自动判定，请把本段输出发给我。")

    print("\n== 财年标签：一致预期的年份覆盖（已报年份应无预期）")
    mar = [c for c in codes if fye.get(c) == 3][:40]
    years = (prev, BASE_YEAR, FORECAST_YEAR)
    cover = {}
    for label, sample in (("12月财年", dec[:60]), ("3月财年", mar)):
        if not sample:
            print(f"   {label}：样本为空")
            continue
        cover[label] = {y: float(west(sample, y).notna().mean()) for y in years}
        print(f"   {label}（{len(sample)} 家）：" + "，".join(f"FY{y} {v:.0%}" for y, v in cover[label].items()))
    cd, cm = cover.get("12月财年"), cover.get("3月财年")
    if cd and cm:
        if cd[prev] > 0.2:
            print(f"   无法判定：12 月财年 FY{prev} 已报仍有预期覆盖，“已报年份无预期”的前提不成立，请把本段输出发给我。")
        elif cm[BASE_YEAR] < 0.2 and cm[FORECAST_YEAR] > 0.5:
            print(f"   -> 约定一致：3 月财年 FY{BASE_YEAR}（截至 {BASE_YEAR} 年 3 月，已于年中报出）无预期，"
                  f"FY{FORECAST_YEAR} 有预期；标签 N = 截至 N 年 3 月，可直接日历化。")
        elif cm[BASE_YEAR] > 0.5:
            print(f"   -> 标签错一年：3 月财年 FY{BASE_YEAR} 仍有预期，说明标签 N 指截至 N+1 年 3 月的财年；"
                  "日历化公式需改为后移一年，请把本段输出发给我。")
        else:
            print("   无法判定：覆盖率处于中间区域，请把本段输出发给我。")
    sample = [c for c in ("9988.HK", "0992.HK") if c in codes]
    if sample:
        print("   样例（原币）：")
        table = pd.DataFrame({f"FY{y}": west(sample, y) for y in range(prev, FORECAST_YEAR + 2)})
        table.insert(0, "财年截止月", fye.reindex(sample))
        print(table.to_string())

    print(f"\n== FY{FORECAST_YEAR + 1} 预期覆盖（日历化需要）")
    nxt = wss(w, CHECK_CODES, "west_netprofit", f"year={FORECAST_YEAR + 1};tradeDate={DATE};westPeriod=180;unit=1", "FY+1")
    print(nxt.to_string())
    no_prof = [c for c in codes if c not in prof.index or pd.isna(prof.get(c))]
    print(f"\n== 财报库覆盖：{len(codes)} 只成分股中，无报表币种识别 {len(no_prof)} 只，"
          f"无财年记录 {int(fye.reindex(codes).isna().sum()) if codes else '—'} 只"
          + (f"（例：{'、'.join(no_prof[:8])}）" if no_prof else ""))
    print("   无币种识别的多为新上市公司：先在财报库跑 fin_update_wind.py hk-estimate / update 补齐。")


# ---------------------------------------------------------------- 主流程
def open_db(C, readonly):
    if readonly:
        return sqlite3.connect(Path(C.DB_PATH).resolve().as_uri() + "?mode=ro", uri=True)
    return sqlite3.connect(C.DB_PATH)


def main(check_only=False, refresh_fx=False):
    from WindPy import w

    C, hk_fx = load_fin_modules()
    core.BASE_YEAR, core.FORECAST_YEAR = BASE_YEAR, FORECAST_YEAR
    core.REVISION_WINDOWS, core.DATE, core.OUT_DIR = REVISION_WINDOWS, DATE, OUT_DIR

    check(w.start().ErrorCode, "Wind 启动")
    if not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")
    con = open_db(C, readonly=not refresh_fx)
    try:
        rate_date = ensure_fx_current(w, con, C, hk_fx, refresh=refresh_fx)
        print(f"财报库：{C.DB_PATH}（{'可写，已检查汇率' if refresh_fx else '只读'}）；中间价日期 {rate_date}", flush=True)
        if check_only:
            run_check(w, con, hk_fx)
            return
        run(w, con, hk_fx, rate_date)
    finally:
        con.close()


def run(w, con, hk_fx, rate_date):
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
    # 基期财年已结束（如 3 月年结的 FY2026 截至 2026-03、已报出）时 Wind 不再保留一致预期，用已报归母净利润（原币）补；
    # 标签 rptDate=BASE_YEAR-12-31 即该财年全年累计（hk_calendar.py 的归期约定）
    base_fy = raw[BASE_YEAR][0]
    fy_end = pd.Series([pd.Timestamp(BASE_YEAR, int(m), 1) + pd.offsets.MonthEnd(0) if m == m else pd.NaT
                        for m in d["财年截止月"]], index=d.index)
    need = base_fy.isna() & (fy_end < pd.Timestamp(DATE))
    if need.any():
        act = wss(w, list(need[need].index), "np_belongto_parcomsh",
                  f"unit=1;rptDate={BASE_YEAR}1231;rptType=1", f"FY{BASE_YEAR} 已报")
        act = pd.to_numeric(act["NP_BELONGTO_PARCOMSH"], errors="coerce").reindex(d.index)
        filled = need & act.notna()
        base_fy = base_fy.where(~filled, act)
        print(f"FY{BASE_YEAR} 已结束的非 12 月财年公司 {int(need.sum())} 只，用已报利润补基期 {int(filled.sum())} 只", flush=True)
    else:
        filled = pd.Series(False, index=d.index)
    d[f"FY{BASE_YEAR}原币"] = base_fy
    d[f"FY{BASE_YEAR}来源"] = np.where(filled, "已报", np.where(base_fy.notna(), "一致预期", ""))
    raw[BASE_YEAR] = (base_fy, raw[BASE_YEAR][1])
    hkd = {y: raw[y][0] * factor for y in fy_years}   # 无系数 -> NaN，自然不计入
    d[np_col(BASE_YEAR)], _ = calendarize(hkd[BASE_YEAR], hkd[FORECAST_YEAR], d["财年截止月"])
    d[np_col(FORECAST_YEAR)], d["口径"] = calendarize(hkd[FORECAST_YEAR], hkd[FORECAST_YEAR + 1], d["财年截止月"])
    for window, past_date in revision_dates.items():
        past = fetch_years(w, codes, (FORECAST_YEAR, FORECAST_YEAR + 1), past_date, f"{window}前")
        # 修正只看预期本身的变化：两个日期用同一个（当前）汇率
        d[np_col(FORECAST_YEAR, window)], _ = calendarize(past[FORECAST_YEAR][0] * factor,
                                                          past[FORECAST_YEAR + 1][0] * factor, d["财年截止月"])
    d["基期提示"] = base_one_off(d)
    d["请求数据日期"] = DATE
    d["成分股返回日期"] = constituent_dates

    done = 0
    for cls_name, cls in CLASSIFICATIONS.items():
        err, ind = w.wss(codes, cls["field"], cls["options"].format(date=DATE.replace("-", "")), usedf=True)
        if err != 0:
            print(f"跳过 {cls_name}：字段 {cls['field']} 返回 ErrorCode={err}", flush=True)
            continue
        ind = normalize_wss(ind, cls_name).reindex(codes).iloc[:, 0]
        # 恒生行业返回值带“(HS)”后缀，如“资讯科技业(HS)”，去掉便于阅读
        labels = ind.astype("string").str.strip().str.replace(r"\s*\(HS\)$", "", regex=True).replace("", pd.NA)
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
                                           ("币种缺失", row["币种未识别市值占比"] > NO_FX_SHARE),
                                           ("基期含一次性", row["基期偏高利润占比"] > ONE_OFF_SHARE)) if hit]
            row["提示"] = "、".join(x for x in [row.get("提示", ""), *extra] if x)
            rows.append(row)
        print(f"{cls_name}：{len(rows)} 个行业（未分类 {int(labels.isna().sum())} 只）", flush=True)
        write_outputs(rows, [dc], hk_notes(revision_dates, len(codes), cls_name, rate_date),
                           f"港股_{cls_name}_{FORECAST_YEAR}盈利预期", cls["title"])
        done += 1
    if not done:
        raise RuntimeError("所有行业分类字段都不可用，请先运行 --check")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="港股分行业一致预期盈利汇总")
    ap.add_argument("--check", action="store_true", help="只做预检：行业字段、一致预期币种、财年标签、财报库覆盖")
    ap.add_argument("--refresh-fx", action="store_true", help="库内中间价未覆盖 DATE 时用 EDB 补齐（会写 fx_parity）")
    args = ap.parse_args()
    main(check_only=args.check, refresh_fx=args.refresh_fx)
