"""申万一级行业：一致预期盈利 + 估值/股息/ROE + 盈利预期修正 + 行业估值历史分位。

在已验证的《申万一级行业一致预期盈利汇总》基础上扩展，原有口径不变：
  * 汇总有一致预期覆盖的上市公司，不代表全行业利润；
  * 两年净利润与增速使用两年均有预测的共同样本；PE 使用市值与利润均非空的同一组股票；
  * 市值、利润按 unit=1 的元口径。

新增：
  1. 更多指标（整体法）：PE_TTM、PB_LF、股息率TTM、ROE_TTM、FY3 预测、营收预期、一致预期 ROE
  2. 盈利预期修正：同一批股票，当前 vs 1个月前 / 3个月前的一致预期净利润（BASE、FORECAST 两个年度），
     含合计修正幅度与修正广度（上调/下调家数占比）
  3. 行业指数估值历史分位（近 N 年 PE_TTM / PB_LF / 股息率）

运行：
  python sw1_consensus_extended.py --check   # 只取 2 个行业前 30 只，核验字段是否有返回，不写文件
  python sw1_consensus_extended.py           # 全量

字段分两类（见 FIELDS_*）：
  * 已验证：原脚本与既有 Wind 取数用过的字段，出错即停止；
  * 待核验：首次使用，出错或全空只告警、该列留空，不影响其他结果。--check 会列出各字段返回情况。
"""
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd

DATE = "2026-11-03"
BASE_YEAR = 2026
FORECAST_YEAR = 2027
EXTRA_YEAR = 2028                 # FY3，覆盖率通常较低，仅作参考
REVISION_LAGS = {"1M": 1, "3M": 3}  # 往前推的月数（按交易日偏移取对应交易日）
REVISION_BAND = 0.01              # 单只股票预测变化超过 ±1% 才计为上调/下调
HISTORY_YEARS = 10                # 行业指数估值分位的回看年数
CHUNK = 800                       # 每次 wss 请求的股票数
WEST_PERIOD = 180
OUT_DIR = Path(os.environ.get("SW1_OUT_DIR", "/Users/amber/Documents/本地数据库/财报"))

SW1 = {
    "801010.SI": "农林牧渔", "801030.SI": "基础化工", "801040.SI": "钢铁",
    "801050.SI": "有色金属", "801080.SI": "电子", "801110.SI": "家用电器",
    "801120.SI": "食品饮料", "801130.SI": "纺织服饰", "801140.SI": "轻工制造",
    "801150.SI": "医药生物", "801160.SI": "公用事业", "801170.SI": "交通运输",
    "801180.SI": "房地产", "801200.SI": "商贸零售", "801210.SI": "社会服务",
    "801230.SI": "综合", "801710.SI": "建筑材料", "801720.SI": "建筑装饰",
    "801730.SI": "电力设备", "801740.SI": "国防军工", "801750.SI": "计算机",
    "801760.SI": "传媒", "801770.SI": "通信", "801780.SI": "银行",
    "801790.SI": "非银金融", "801880.SI": "汽车", "801890.SI": "机械设备",
    "801950.SI": "煤炭", "801960.SI": "石油石化", "801970.SI": "环保",
    "801980.SI": "美容护理",
}

# 已验证字段（原脚本 / 既有 Wind 取数已使用）
FIELDS_MKT = ["sec_name", "mkt_cap_ard"]
FIELDS_VAL = ["pe_ttm", "pb_lf", "dividendyield2"]          # 既有 HSHDYI.HI 等取数已用
FIELDS_WEST = ["west_netprofit", "west_instnum"]
# 待核验字段：逐个单独请求，失败只告警
FIELDS_WEST_OPTIONAL = {"west_sales": "SALES", "west_avgroe": "ROE_E"}
FIELDS_MKT_OPTIONAL = {"mkt_cap_float": "MKT_CAP_FLOAT"}


# ---------------------------------------------------------------- 通用工具（沿用原脚本）
def check(err, what):
    if err != 0:
        raise RuntimeError(f"{what} 失败，ErrorCode={err}")


def require_columns(frame, columns, what):
    missing = set(columns) - set(frame.columns)
    if missing:
        raise RuntimeError(f"{what} 缺少字段：{sorted(missing)}")


def normalize_wss(frame, what):
    frame = frame.copy()
    frame.columns = frame.columns.astype(str).str.upper()
    frame.index = frame.index.astype(str).str.strip().str.upper()
    if not frame.index.is_unique:
        raise RuntimeError(f"{what} 返回重复股票代码")
    return frame


def ratio(numerator, denominator):
    """分母不为正时留空。"""
    if pd.isna(numerator) or pd.isna(denominator) or denominator <= 0:
        return float("nan")
    return numerator / denominator


def chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


# ---------------------------------------------------------------- 取数
def wss_frame(w, codes, fields, options, what, strict=True):
    """分批 wss；strict=False 时失败返回全空列并告警。"""
    parts = []
    for part in chunks(codes, CHUNK):
        err, df = w.wss(part, ",".join(fields), options, usedf=True)
        if err != 0:
            if strict:
                check(err, what)
            print(f"警告：{what} 返回 ErrorCode={err}，该字段留空", flush=True)
            return pd.DataFrame(index=codes, columns=[f.upper() for f in fields], dtype=float)
        parts.append(normalize_wss(df, what))
    out = pd.concat(parts)
    if not out.index.is_unique:
        raise RuntimeError(f"{what} 分批合并后出现重复代码")
    out = out.reindex(codes)
    require_columns(out, [f.upper() for f in fields], what)
    return out


def trade_date_offset(w, months):
    """DATE 往前推 months 个月对应的交易日。"""
    res = w.tdaysoffset(-months, DATE, "Period=M")
    check(res.ErrorCode, f"tdaysoffset -{months}M")
    return pd.Timestamp(res.Data[0][0]).strftime("%Y-%m-%d")


def get_constituents(w, idx, name):
    err, cons = w.wset("indexconstituent", f"date={DATE};windcode={idx}", usedf=True)
    check(err, f"{name} 成分股")
    cons = cons.copy()
    cons.columns = cons.columns.astype(str).str.lower()
    require_columns(cons, ["wind_code"], f"{name} 成分股")
    codes = cons["wind_code"].dropna().astype(str).str.strip().str.upper()
    codes = codes[codes.ne("")].drop_duplicates().tolist()
    if not codes:
        raise RuntimeError(f"{name} 成分股为空，停止导出")
    returned = "未返回日期"
    if "date" in cons.columns:
        dates = pd.to_datetime(cons["date"], errors="coerce").dropna()
        returned = ",".join(sorted(set(dates.dt.strftime("%Y-%m-%d"))))
        if returned != DATE:
            print(f"警告：{name} 请求日期 {DATE}，成分股返回日期 {returned}", flush=True)
    return codes, returned


def fetch_stock_panel(w, codes, rev_dates):
    """全体成分股一次性取数（各行业共用），返回个股宽表。"""
    d = wss_frame(w, codes, FIELDS_MKT + FIELDS_VAL, f"tradeDate={DATE};unit=1", "市值与估值")
    for c in ["MKT_CAP_ARD"] + [f.upper() for f in FIELDS_VAL]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    if d["MKT_CAP_ARD"].isna().all():
        raise RuntimeError(f"{DATE} 市值全部为空，请检查数据日期；未导出本次结果")
    for f, col in FIELDS_MKT_OPTIONAL.items():
        e = wss_frame(w, codes, [f], f"tradeDate={DATE};unit=1", f"{f}(待核验)", strict=False)
        d[col] = pd.to_numeric(e[f.upper()], errors="coerce")

    years = [BASE_YEAR, FORECAST_YEAR, EXTRA_YEAR]
    for year in years:
        e = wss_frame(w, codes, FIELDS_WEST,
                      f"year={year};tradeDate={DATE};westPeriod={WEST_PERIOD};unit=1", f"{year} 预期")
        d[f"NP{year}"] = pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce")
        d[f"INST{year}"] = pd.to_numeric(e["WEST_INSTNUM"], errors="coerce")
        if d[f"NP{year}"].isna().all():
            print(f"警告：{year} 年预测利润全部为空", flush=True)
        if year == EXTRA_YEAR:
            continue
        for f, col in FIELDS_WEST_OPTIONAL.items():
            e = wss_frame(w, codes, [f], f"year={year};tradeDate={DATE};westPeriod={WEST_PERIOD};unit=1",
                          f"{year} {f}(待核验)", strict=False)
            d[f"{col}{year}"] = pd.to_numeric(e[f.upper()], errors="coerce")

    # 盈利预期修正：同一批股票在过去日期的一致预期
    for lag, rdate in rev_dates.items():
        for year in (BASE_YEAR, FORECAST_YEAR):
            e = wss_frame(w, codes, ["west_netprofit"],
                          f"year={year};tradeDate={rdate};westPeriod={WEST_PERIOD};unit=1", f"{year} 预期@{rdate}")
            d[f"NP{year}_{lag}"] = pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce")
    return d


def fetch_index_history(w, start):
    """行业指数估值历史（日频），返回 {字段: DataFrame(日期 x 指数)}。"""
    out = {}
    for f in FIELDS_VAL:
        err, df = w.wsd(list(SW1), f, start, DATE, "", usedf=True)
        if err != 0:
            print(f"警告：行业指数 {f} 历史 ErrorCode={err}，估值分位留空", flush=True)
            continue
        df = df.copy()
        df.index = pd.to_datetime(df.index)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        out[f] = df.apply(pd.to_numeric, errors="coerce")
    return out


# ---------------------------------------------------------------- 汇总
def summarize(d, name):
    """整体法。PE/PB/ROE 分母为正才计算；股息率按市值加权（=合计股息/合计市值）。"""
    mv = d["MKT_CAP_ARD"]
    total_mv = mv.sum(min_count=1)
    nb, nf, nx = f"NP{BASE_YEAR}", f"NP{FORECAST_YEAR}", f"NP{EXTRA_YEAR}"
    both = d.dropna(subset=[nb, nf])
    has_f = d.dropna(subset=[nf])
    pe_f = d.dropna(subset=["MKT_CAP_ARD", nf])
    pe_b = d.dropna(subset=["MKT_CAP_ARD", nb])
    three = d.dropna(subset=[nf, nx])

    # TTM 利润与净资产由市值/PE、市值/PB 反推（PE 为负 → 亏损；PB<=0 → 净资产为负，剔除）
    ttm = d.dropna(subset=["MKT_CAP_ARD", "PE_TTM"])
    ttm = ttm[ttm["PE_TTM"] != 0]
    eq = d.dropna(subset=["MKT_CAP_ARD", "PB_LF"])
    eq = eq[eq["PB_LF"] > 0]
    equity = eq["MKT_CAP_ARD"] / eq["PB_LF"]
    roe_s = d.loc[eq.index].dropna(subset=["PE_TTM"])
    roe_s = roe_s[roe_s["PE_TTM"] != 0]
    dy = d.dropna(subset=["MKT_CAP_ARD", "DIVIDENDYIELD2"])
    roe_f = d.loc[eq.index].dropna(subset=[nf])

    row = {
        "行业": name,
        "成分股数": len(d),
        "总市值(亿)": total_mv / 1e8,
        "市值缺失家数": int(mv.isna().sum()),
        f"有{FORECAST_YEAR}预测家数": len(has_f),
        "共同样本家数": len(both),
        "预测样本市值覆盖率(已知市值口径)": ratio(has_f["MKT_CAP_ARD"].sum(min_count=1), total_mv),
        "共同样本市值覆盖率(已知市值口径)": ratio(both["MKT_CAP_ARD"].sum(min_count=1), total_mv),
        f"{BASE_YEAR}E净利润(亿,共同样本)": both[nb].sum(min_count=1) / 1e8,
        f"{FORECAST_YEAR}E净利润(亿,共同样本)": both[nf].sum(min_count=1) / 1e8,
        f"{FORECAST_YEAR}E增速(共同样本)": ratio(both[nf].sum(min_count=1), both[nb].sum(min_count=1)) - 1,
        f"{EXTRA_YEAR}E增速(与{FORECAST_YEAR}共同样本)": ratio(three[nx].sum(min_count=1), three[nf].sum(min_count=1)) - 1,
        f"{EXTRA_YEAR}共同样本家数": len(three),
        f"{BASE_YEAR}E PE": ratio(pe_b["MKT_CAP_ARD"].sum(min_count=1), pe_b[nb].sum(min_count=1)),
        f"{FORECAST_YEAR}E PE(市值利润齐全样本)": ratio(pe_f["MKT_CAP_ARD"].sum(min_count=1), pe_f[nf].sum(min_count=1)),
        "PE_TTM(整体法)": ratio(ttm["MKT_CAP_ARD"].sum(min_count=1), (ttm["MKT_CAP_ARD"] / ttm["PE_TTM"]).sum(min_count=1)),
        "PB_LF(整体法,剔除负净资产)": ratio(eq["MKT_CAP_ARD"].sum(min_count=1), equity.sum(min_count=1)),
        "ROE_TTM(整体法)": ratio((roe_s["MKT_CAP_ARD"] / roe_s["PE_TTM"]).sum(min_count=1),
                                (roe_s["MKT_CAP_ARD"] / roe_s["PB_LF"]).sum(min_count=1)),
        f"{FORECAST_YEAR}E ROE(预测利润/当前净资产)": ratio(roe_f[nf].sum(min_count=1),
                                                     (roe_f["MKT_CAP_ARD"] / roe_f["PB_LF"]).sum(min_count=1)),
        "股息率TTM(市值加权,%)": ratio((dy["MKT_CAP_ARD"] * dy["DIVIDENDYIELD2"]).sum(min_count=1),
                                    dy["MKT_CAP_ARD"].sum(min_count=1)),
        f"平均机构数({FORECAST_YEAR}预测样本)": has_f[f"INST{FORECAST_YEAR}"].mean(),
    }
    # 营收与一致预期 ROE（待核验字段；全空则留空）
    s_b, s_f = f"SALES{BASE_YEAR}", f"SALES{FORECAST_YEAR}"
    if s_b in d and s_f in d:
        sb = d.dropna(subset=[s_b, s_f])
        row[f"{FORECAST_YEAR}E营收增速(共同样本)"] = ratio(sb[s_f].sum(min_count=1), sb[s_b].sum(min_count=1)) - 1
        sm = d.dropna(subset=[s_f, nf])
        row[f"{FORECAST_YEAR}E净利率(共同样本)"] = ratio(sm[nf].sum(min_count=1), sm[s_f].sum(min_count=1))
    r_f = f"ROE_E{FORECAST_YEAR}"
    if r_f in d:
        rr = d.dropna(subset=[r_f, "MKT_CAP_ARD"])
        row[f"{FORECAST_YEAR}E一致预期ROE均值%(市值加权)"] = ratio((rr[r_f] * rr["MKT_CAP_ARD"]).sum(min_count=1),
                                                             rr["MKT_CAP_ARD"].sum(min_count=1))
    return row


def revisions(d, name, rev_dates):
    """同一批股票（两个时点均有预测）的合计修正幅度和修正广度。"""
    row = {"行业": name}
    for lag, rdate in rev_dates.items():
        for year in (BASE_YEAR, FORECAST_YEAR):
            now, past = f"NP{year}", f"NP{year}_{lag}"
            s = d.dropna(subset=[now, past])
            row[f"{year}E 样本家数({lag})"] = len(s)
            row[f"{year}E 合计修正({lag})"] = ratio(s[now].sum(min_count=1), s[past].sum(min_count=1)) - 1
            pos = s[s[past] > 0]
            chg = pos[now] / pos[past] - 1
            n = len(pos)
            row[f"{year}E 上调占比({lag})"] = (chg > REVISION_BAND).mean() if n else np.nan
            row[f"{year}E 下调占比({lag})"] = (chg < -REVISION_BAND).mean() if n else np.nan
            row[f"{year}E 修正广度({lag})"] = row[f"{year}E 上调占比({lag})"] - row[f"{year}E 下调占比({lag})"] if n else np.nan
        row[f"{lag}基准日"] = rdate
    return row


def percentile_table(hist):
    rows = []
    label = {"pe_ttm": "PE_TTM", "pb_lf": "PB_LF", "dividendyield2": "股息率TTM"}
    for idx, name in SW1.items():
        row = {"行业": name, "指数代码": idx}
        for f, df in hist.items():
            col = idx.upper()
            if col not in df:
                continue
            s = df[col].dropna()
            if f == "pe_ttm":
                s = s[s > 0]  # 负 PE 不参与分位
            if s.empty:
                continue
            cur = s.iloc[-1]
            row[f"{label[f]}当前"] = cur
            row[f"{label[f]}分位"] = (s <= cur).mean()
            row[f"{label[f]}中位数"] = s.median()
            row[f"{label[f]}起始日"] = s.index[0].strftime("%Y-%m-%d")
        rows.append(row)
    return pd.DataFrame(rows)


def field_coverage(d):
    cols = [c for c in d.columns if c not in ("SEC_NAME", "行业", "请求数据日期", "成分股返回日期")]
    return pd.DataFrame({"字段": cols, "非空家数": [int(d[c].notna().sum()) for c in cols],
                         "总家数": len(d)}).assign(覆盖率=lambda x: x["非空家数"] / x["总家数"])


# ---------------------------------------------------------------- 主流程
def main(check_only=False):
    from WindPy import w

    check(w.start().ErrorCode, "Wind 启动")
    if not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")

    industries = SW1
    if check_only:
        industries = {k: v for k, v in SW1.items() if v in ("银行", "电子")}

    rev_dates = {lag: trade_date_offset(w, m) for lag, m in REVISION_LAGS.items()}
    print(f"修正基准日：{rev_dates}", flush=True)

    members, cons_dates = {}, {}
    for idx, name in industries.items():
        codes, returned = get_constituents(w, idx, name)
        members[name] = codes[:30] if check_only else codes
        cons_dates[name] = returned
    all_codes = list(dict.fromkeys(c for cs in members.values() for c in cs))
    print(f"成分股合计 {len(all_codes)} 只（去重）", flush=True)

    panel = fetch_stock_panel(w, all_codes, rev_dates)
    if check_only:
        print(field_coverage(panel).to_string(index=False))
        print("核验模式：未写文件。覆盖率为 0 的待核验字段请在终端确认字段名或参数。")
        return

    rows, revs, details = [], [], []
    for name, codes in members.items():
        d = panel.loc[codes].copy()
        if d["MKT_CAP_ARD"].isna().any():
            print(f"警告：{name} 有 {d['MKT_CAP_ARD'].isna().sum()} 只股票缺失市值", flush=True)
        rows.append(summarize(d, name))
        revs.append(revisions(d, name, rev_dates))
        d["行业"] = name
        d["请求数据日期"] = DATE
        d["成分股返回日期"] = cons_dates[name]
        details.append(d)
        print(f"完成 {name}: {len(codes)} 只", flush=True)

    summary = pd.DataFrame(rows).sort_values(f"{FORECAST_YEAR}E增速(共同样本)", ascending=False, na_position="last")
    rev = pd.DataFrame(revs).sort_values(f"{FORECAST_YEAR}E 合计修正(1M)", ascending=False, na_position="last")
    start = (pd.Timestamp(DATE) - pd.DateOffset(years=HISTORY_YEARS)).strftime("%Y-%m-%d")
    pct = percentile_table(fetch_index_history(w, start))
    detail = pd.concat(details).rename_axis("股票代码")

    notes = pd.DataFrame({"统计说明": [
        f"请求数据日期：{DATE}；成分股实际返回日期见个股明细。修正基准日：{rev_dates}。",
        "本表汇总有 Wind 一致预期覆盖的成分股，不代表全行业盈利。",
        "两年净利润与增速使用两年均有预测的共同样本；FY3 增速使用与 FY2 的共同样本，覆盖率较低，仅作参考。",
        "基期合计利润小于等于零时不展示常规增速；预测合计利润小于等于零时不展示PE。",
        "PE_TTM、ROE_TTM 由 市值/PE_TTM 反推 TTM 利润后整体法合计；PB 由 市值/PB_LF 反推净资产，剔除净资产为负的公司。",
        f"{FORECAST_YEAR}E ROE = 预测净利润合计 / 当前净资产合计（近似，未考虑留存增厚）。",
        "股息率TTM为 dividendyield2 按市值加权，等于合计股息/合计市值。",
        "盈利修正：以当前成分股为样本，比较当前与基准日的一致预期净利润；只用两个时点均有预测的股票。",
        f"修正广度 = 上调占比 − 下调占比，单只股票变化超过 ±{REVISION_BAND:.0%} 才计入；基准日预测利润为正的股票才计入广度。",
        "合计修正受大市值公司主导，广度反映家数分布，两者需一起看。",
        f"估值分位：申万一级指数近 {HISTORY_YEARS} 年日频，当前值在历史中的位置（≤当前的天数占比）；负 PE 不参与。",
        "待核验字段（mkt_cap_float、west_sales、west_avgroe）失败或全空时相关列留空，不影响其他结果。",
        "市值和利润按unit=1的元口径处理，利润除以1e8换算为亿元；正式使用前核对终端定义。",
        "机构数沿用west_instnum及year参数，年度参数的适用性仍需在终端字段说明中核对。",
    ]})

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_file = OUT_DIR / f"申万一级_盈利预期与修正_{DATE}.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="行业汇总", index=False)
        rev.to_excel(xw, sheet_name="盈利修正", index=False)
        pct.to_excel(xw, sheet_name="估值分位", index=False)
        detail.to_excel(xw, sheet_name="个股明细")
        field_coverage(detail).to_excel(xw, sheet_name="字段覆盖", index=False)
        notes.to_excel(xw, sheet_name="统计说明", index=False)
        for sheet, frame in (("行业汇总", summary), ("盈利修正", rev), ("估值分位", pct)):
            ws = xw.sheets[sheet]
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            for column, title in enumerate(frame.columns, start=1):
                if any(k in title for k in ("覆盖率", "增速", "修正", "占比", "广度", "分位", "ROE(", "ROE_TTM", "净利率")):
                    for cells in ws.iter_rows(min_row=2, min_col=column, max_col=column):
                        cells[0].number_format = "0.00%"
    print(summary.round(3).to_string(index=False))
    print(rev.round(3).to_string(index=False))
    print(f"已保存：{out_file}")


if __name__ == "__main__":
    main(check_only="--check" in sys.argv)
