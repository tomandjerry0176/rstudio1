"""申万一级行业一致预期盈利汇总（含预期修正与集中度）。

运行环境需要 WindPy、pandas、openpyxl，并已登录可用的 Wind 终端。
DATE 默认为已验证有市值数据的交易日；不自动使用尚未更新完整的盘中数据。
本脚本汇总有一致预期覆盖的上市公司，不代表全行业利润。
市值及预测利润按 unit=1 的元口径处理，正式使用前仍需核对终端字段定义。
"""
from pathlib import Path

import numpy as np
import pandas as pd

DATE = "2026-11-03"
BASE_YEAR = 2026
FORECAST_YEAR = 2027
# 预期修正回看窗口：标签 -> 向前偏移的交易日数（约 4 周 / 13 周）
REVISION_WINDOWS = {"4周": 20, "13周": 65}
TOP_N = 3
# 提示标签阈值
SMALL_SAMPLE = 10          # 有预测家数少于此值 -> 小样本
TURNAROUND_SHARE = 0.5     # 基期亏损公司贡献的增量占比超过此值 -> 扭亏驱动
SINGLE_NAME_SHARE = 0.35   # 单一公司贡献的增量占比超过此值 -> 单一公司主导
CONCENTRATION = 0.6        # 前 TOP_N 公司利润占比超过此值 -> 集中度高
OUT_DIR = Path("/Users/amber/Documents/本地数据库/财报")

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


def positive_denominator_ratio(numerator, denominator):
    if pd.isna(numerator) or pd.isna(denominator) or denominator <= 0:
        return float("nan")
    return numerator / denominator


def np_col(year, window=None):
    return f"NP{year}" if window is None else f"NP{year}_{window}前"


def offset_trade_date(w, days):
    """DATE 向前 days 个交易日（Wind 交易日历）。"""
    data = w.tdaysoffset(-days, DATE)
    check(data.ErrorCode, f"交易日偏移 -{days}")
    return pd.Timestamp(data.Data[0][0]).strftime("%Y-%m-%d")


def fetch_forecast(w, codes, year, trade_date, what):
    err, e = w.wss(
        codes, "west_netprofit,west_instnum",
        f"year={year};tradeDate={trade_date};westPeriod=180;unit=1", usedf=True,
    )
    check(err, what)
    e = normalize_wss(e, what).reindex(codes)
    require_columns(e, ["WEST_NETPROFIT", "WEST_INSTNUM"], what)
    return (pd.to_numeric(e["WEST_NETPROFIT"], errors="coerce"),
            pd.to_numeric(e["WEST_INSTNUM"], errors="coerce"))


def revision_stats(d, window):
    """同一批股票当前与 window 前的 FORECAST_YEAR 预测利润对比。"""
    now, past = np_col(FORECAST_YEAR), np_col(FORECAST_YEAR, window)
    s = d.dropna(subset=[now, past])
    now_sum, past_sum = s[now].sum(min_count=1), s[past].sum(min_count=1)
    # 当前或对比日合计利润小于等于零时，比率没有意义（如房地产由微利转亏会得到 -2000%），只看修正额
    change = (positive_denominator_ratio(now_sum, past_sum)
              if pd.notna(now_sum) and now_sum > 0 else float("nan"))
    same = np.isclose(s[now], s[past], rtol=1e-6, atol=0)
    up = int(((s[now] > s[past]) & ~same).sum())
    down = int(((s[now] < s[past]) & ~same).sum())
    return {
        f"{FORECAST_YEAR}E修正_{window}(配对样本)": change - 1,
        f"{FORECAST_YEAR}E修正额_{window}(亿)": (now_sum - past_sum) / 1e8,
        f"修正样本家数_{window}": len(s),
        f"上调家数_{window}": up,
        f"下调家数_{window}": down,
        f"上调净占比_{window}": (up - down) / len(s) if len(s) else float("nan"),
    }


def concentration_stats(d):
    """FORECAST_YEAR 预测利润前 TOP_N 公司占比与个股增速中位数。"""
    np_base, np_forecast = np_col(BASE_YEAR), np_col(FORECAST_YEAR)
    has_forecast = d.dropna(subset=[np_forecast])
    total = has_forecast[np_forecast].sum(min_count=1)
    top = has_forecast.nlargest(TOP_N, np_forecast)
    pos_base = d.dropna(subset=[np_base, np_forecast])
    pos_base = pos_base[pos_base[np_base] > 0]
    stock_growth = pos_base[np_forecast] / pos_base[np_base] - 1
    return {
        f"前{TOP_N}公司{FORECAST_YEAR}E利润占比": positive_denominator_ratio(
            top[np_forecast].sum(min_count=1), total
        ),
        f"前{TOP_N}公司": "、".join(top["SEC_NAME"].fillna(pd.Series(top.index, index=top.index))),
        f"个股{FORECAST_YEAR}E增速中位数(基期利润>0)": (
            stock_growth.median() if len(stock_growth) else float("nan")
        ),
    }


def increment_stats(d):
    """共同样本内的增量来源：扭亏贡献、最大增量公司及剔除后增速。"""
    np_base, np_forecast = np_col(BASE_YEAR), np_col(FORECAST_YEAR)
    both = d.dropna(subset=[np_base, np_forecast])
    delta = both[np_forecast] - both[np_base]
    total_delta = delta.sum(min_count=1)
    turnaround = delta[both[np_base] <= 0].sum()
    out = {
        f"{FORECAST_YEAR}E增量(亿,共同样本)": total_delta / 1e8,
        "扭亏贡献占比": positive_denominator_ratio(turnaround, total_delta),
        "最大增量公司": "",
        "最大增量公司贡献占比": float("nan"),
        "剔除最大增量公司后增速": float("nan"),
    }
    if len(delta):
        code = delta.idxmax()
        out["最大增量公司"] = both.at[code, "SEC_NAME"] if pd.notna(both.at[code, "SEC_NAME"]) else code
        out["最大增量公司贡献占比"] = positive_denominator_ratio(delta[code], total_delta)
        rest = both.drop(index=code)
        out["剔除最大增量公司后增速"] = positive_denominator_ratio(
            rest[np_forecast].sum(min_count=1), rest[np_base].sum(min_count=1)
        ) - 1
    return out


def flag_text(row):
    flags = []
    if row[f"有{FORECAST_YEAR}预测家数"] < SMALL_SAMPLE:
        flags.append("小样本")
    base = row[f"{BASE_YEAR}E净利润(亿,共同样本)"]
    if pd.notna(base) and base <= 0:
        flags.append("基期亏损")
    elif row["扭亏贡献占比"] > TURNAROUND_SHARE:
        flags.append("扭亏驱动")
    if row["最大增量公司贡献占比"] > SINGLE_NAME_SHARE:
        flags.append("单一公司主导")
    if row[f"前{TOP_N}公司{FORECAST_YEAR}E利润占比"] > CONCENTRATION:
        flags.append("集中度高")
    return "、".join(flags)


def summarize(d, name):
    """无效比率留空；PE 的市值和利润严格使用同一组股票。"""
    np_base, np_forecast = np_col(BASE_YEAR), np_col(FORECAST_YEAR)
    both = d.dropna(subset=[np_base, np_forecast])
    has_forecast = d.dropna(subset=[np_forecast])
    pe_sample = d.dropna(subset=["MKT_CAP_ARD", np_forecast])
    total_mv = d["MKT_CAP_ARD"].sum(min_count=1)
    base_profit = both[np_base].sum(min_count=1)
    forecast_profit = both[np_forecast].sum(min_count=1)
    growth_ratio = positive_denominator_ratio(forecast_profit, base_profit)
    row = {
        "行业": name,
        "成分股数": len(d),
        "市值缺失家数": int(d["MKT_CAP_ARD"].isna().sum()),
        f"有{FORECAST_YEAR}预测家数": len(has_forecast),
        "共同样本家数": len(both),
        "PE样本家数": len(pe_sample),
        "预测样本市值覆盖率(已知市值口径)": positive_denominator_ratio(
            has_forecast["MKT_CAP_ARD"].sum(min_count=1), total_mv
        ),
        "共同样本市值覆盖率(已知市值口径)": positive_denominator_ratio(
            both["MKT_CAP_ARD"].sum(min_count=1), total_mv
        ),
        f"{BASE_YEAR}E净利润(亿,共同样本)": base_profit / 1e8,
        f"{FORECAST_YEAR}E净利润(亿,共同样本)": forecast_profit / 1e8,
        f"{FORECAST_YEAR}E增速(共同样本)": growth_ratio - 1,
        f"{FORECAST_YEAR}E PE(市值利润齐全样本)": positive_denominator_ratio(
            pe_sample["MKT_CAP_ARD"].sum(min_count=1),
            pe_sample[np_forecast].sum(min_count=1),
        ),
        f"平均机构数({FORECAST_YEAR}预测样本)": has_forecast[f"INST{FORECAST_YEAR}"].mean(),
    }
    row.update(increment_stats(d))
    row.update(concentration_stats(d))
    for window in REVISION_WINDOWS:
        row.update(revision_stats(d, window))
    row["提示"] = flag_text(row)
    return row


def note_lines(revision_dates, date=None):
    """统计说明文字；看板生成时也调用它，保证说明与当前口径一致。"""
    date = date or DATE
    window_text = "；".join(f"{win}前={dt}" for win, dt in revision_dates.items())
    return [
        f"请求数据日期：{date}；成分股实际返回日期见个股明细。",
        "本表汇总有 Wind 一致预期覆盖的成分股，不代表全行业盈利。",
        "两年净利润与增速使用两年均有预测的共同样本。",
        "基期合计利润小于等于零时不展示常规增速；预测合计利润小于等于零时不展示PE。",
        "PE使用预测利润与市值均非空的同一组股票，保留样本中的亏损公司。",
        "覆盖率分母仅为已取得的成分股市值；存在市值缺失时不是完整行业市值覆盖率。",
        "市值和利润按unit=1的元口径处理，利润除以1e8换算为亿元；正式使用前核对终端定义。",
        "机构数沿用west_instnum及year参数，年度参数的适用性仍需在终端字段说明中核对。",
        f"预期修正对比日期（交易日偏移）：{window_text}。",
        f"预期修正 = 当前与对比日均有{FORECAST_YEAR}预测的同一批股票，预测利润合计之比减1；"
        "当前或对比日合计小于等于零时比率留空，此时参考修正额（亿元）。",
        "预期修正使用当前成分股回看历史预期，不重建历史成分股；新覆盖的公司不计入修正。",
        "上调/下调家数按个股预测利润变化判断，相对变动小于百万分之一视为不变；"
        "上调净占比=(上调-下调)/修正样本家数。",
        f"前{TOP_N}公司占比 = 预测利润最大的{TOP_N}家之和 / 全部有预测公司合计；"
        "合计小于等于零时留空，样本含亏损公司时可能超过100%。",
        f"个股增速中位数仅统计基期预测利润大于零的公司，用于对照受大公司主导的合计增速。",
        f"增量 = 共同样本{FORECAST_YEAR}E与{BASE_YEAR}E预测利润之差；扭亏贡献占比 = 基期预测利润"
        "小于等于零的公司贡献的增量 / 行业总增量；总增量小于等于零时留空。",
        "最大增量公司 = 共同样本中增量最大的一家；剔除后增速为去掉该公司后的共同样本增速。",
        f"提示标签：有预测家数<{SMALL_SAMPLE}为小样本；基期合计亏损为基期亏损；扭亏贡献>"
        f"{TURNAROUND_SHARE:.0%}为扭亏驱动；单一公司增量占比>{SINGLE_NAME_SHARE:.0%}为单一公司主导；"
        f"前{TOP_N}公司利润占比>{CONCENTRATION:.0%}为集中度高。",
        "Excel空白表示数据缺失或该指标不适用；市值全部为空时终止，不写入本次结果。",
    ]


def main():
    from WindPy import w

    check(w.start().ErrorCode, "Wind 启动")
    if not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")

    revision_dates = {win: offset_trade_date(w, days) for win, days in REVISION_WINDOWS.items()}
    print(f"预期修正对比日期：{revision_dates}", flush=True)

    rows, details = [], []
    for idx, name in SW1.items():
        err, cons = w.wset("indexconstituent", f"date={DATE};windcode={idx}", usedf=True)
        check(err, f"{name} 成分股")
        cons = cons.copy()
        cons.columns = cons.columns.astype(str).str.lower()
        require_columns(cons, ["wind_code"], f"{name} 成分股")
        codes = cons["wind_code"].dropna().astype(str).str.strip().str.upper()
        codes = codes[codes.ne("")].drop_duplicates().tolist()
        if not codes:
            raise RuntimeError(f"{name} 成分股为空，停止导出")
        constituent_dates = "未返回日期"
        if "date" in cons.columns:
            dates = pd.to_datetime(cons["date"], errors="coerce").dropna()
            constituent_dates = ",".join(sorted(set(dates.dt.strftime("%Y-%m-%d"))))
            if constituent_dates != DATE:
                print(f"警告：{name} 请求日期 {DATE}，成分股返回日期 {constituent_dates}", flush=True)

        err, d = w.wss(codes, "sec_name,mkt_cap_ard", f"tradeDate={DATE};unit=1", usedf=True)
        check(err, f"{name} 市值")
        d = normalize_wss(d, f"{name} 市值").reindex(codes)
        require_columns(d, ["SEC_NAME", "MKT_CAP_ARD"], f"{name} 市值")
        d["MKT_CAP_ARD"] = pd.to_numeric(d["MKT_CAP_ARD"], errors="coerce")
        if d["MKT_CAP_ARD"].isna().all():
            raise RuntimeError(f"{name} {DATE} 市值全部为空，请检查数据日期；未导出本次结果")
        if d["MKT_CAP_ARD"].isna().any():
            print(f"警告：{name} 有 {d['MKT_CAP_ARD'].isna().sum()} 只股票缺失市值", flush=True)

        for year in (BASE_YEAR, FORECAST_YEAR):
            d[np_col(year)], d[f"INST{year}"] = fetch_forecast(
                w, codes, year, DATE, f"{name} {year} 预期"
            )
            if d[np_col(year)].isna().all():
                print(f"警告：{name} {year} 年预测利润全部为空", flush=True)
        # 预期修正：用当前成分股回看历史一致预期，样本口径与当前一致
        for window, past_date in revision_dates.items():
            d[np_col(FORECAST_YEAR, window)], _ = fetch_forecast(
                w, codes, FORECAST_YEAR, past_date, f"{name} {FORECAST_YEAR} 预期({window}前)"
            )
        d["行业"] = name
        d["请求数据日期"] = DATE
        d["成分股返回日期"] = constituent_dates
        rows.append(summarize(d, name))
        details.append(d)
        print(f"完成 {name}: {len(codes)} 只", flush=True)

    summary = pd.DataFrame(rows).sort_values(
        f"{FORECAST_YEAR}E增速(共同样本)", ascending=False, na_position="last"
    )
    notes = pd.DataFrame({"统计说明": note_lines(revision_dates)})
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_file = OUT_DIR / f"申万一级_{FORECAST_YEAR}盈利预期_{DATE}.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="行业汇总", index=False)
        pd.concat(details).rename_axis("股票代码").to_excel(xw, sheet_name="个股明细")
        notes.to_excel(xw, sheet_name="统计说明", index=False)
        ws = xw.sheets["行业汇总"]
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        percent_keys = ("覆盖率", "增速", "修正_", "占比", "贡献占比")
        for column, title in enumerate(summary.columns, start=1):
            if any(key in title for key in percent_keys):
                for cells in ws.iter_rows(min_row=2, min_col=column, max_col=column):
                    cells[0].number_format = "0.00%"
    print(summary.round(3).to_string(index=False))
    print(f"已保存：{out_file}")

    from build_dashboard import build
    for html_file in (out_file.with_suffix(".html"),
                      OUT_DIR / f"申万一级_{FORECAST_YEAR}盈利预期_latest.html"):
        build(out_file, html_file)
        print(f"看板已保存：{html_file}")


if __name__ == "__main__":
    main()
