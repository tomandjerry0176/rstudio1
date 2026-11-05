"""由 sw1_consensus_2027.py 导出的 Excel 生成单文件 HTML 看板。

用法：python build_dashboard.py 输入.xlsx [输出.html] [--title 标题]
指标从“个股明细”重新计算（与主脚本同一套 summarize 逻辑），旧版 Excel 也能生成新指标。
看板为离线单文件，不依赖任何外部脚本或字体，可直接放进本地工作台的 iframe。
"""
import json
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import sw1_consensus_2027 as core

TOP_STOCKS = 10
TEMPLATE = Path(__file__).with_name("dashboard_template.html")


def clean(value):
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (float, np.floating)):
        return None if not np.isfinite(value) else round(float(value), 6)
    if value is None or (isinstance(value, str) and value == ""):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def ratio(num, den):
    return core.positive_denominator_ratio(num, den)


def configure_core(details):
    years = sorted(int(c[2:]) for c in details.columns if re.fullmatch(r"NP\d{4}", c))
    if len(years) < 2:
        raise RuntimeError("个股明细缺少两年预测利润列（NPyyyy）")
    core.BASE_YEAR, core.FORECAST_YEAR = years[0], years[-1]
    pattern = re.compile(rf"NP{core.FORECAST_YEAR}_(.+)前$")
    windows = [m.group(1) for c in details.columns if (m := pattern.match(c))]
    core.REVISION_WINDOWS = {w: None for w in windows}
    return windows


def overall_stats(d, windows):
    b, f = core.np_col(core.BASE_YEAR), core.np_col(core.FORECAST_YEAR)
    both = d.dropna(subset=[b, f])
    pe = d.dropna(subset=["MKT_CAP_ARD", f])
    out = {
        "stocks": len(d),
        "nf": int(d[f].notna().sum()),
        "np0": both[b].sum() / 1e8,
        "np1": both[f].sum() / 1e8,
        "g": ratio(both[f].sum(), both[b].sum()) - 1,
        "pe": ratio(pe["MKT_CAP_ARD"].sum(), pe[f].sum()),
        "rev": {},
    }
    for w in windows:
        s = d.dropna(subset=[f, core.np_col(core.FORECAST_YEAR, w)])
        past = s[core.np_col(core.FORECAST_YEAR, w)].sum()
        out["rev"][w] = ratio(s[f].sum(), past) - 1 if s[f].sum() > 0 else None
    return out


def stock_rows(g, windows):
    b, f = core.np_col(core.BASE_YEAR), core.np_col(core.FORECAST_YEAR)
    both = g.dropna(subset=[b, f]).copy()
    both["delta"] = both[f] - both[b]
    both = both.reindex(both["delta"].abs().sort_values(ascending=False).index).head(TOP_STOCKS)
    rows = []
    for code, r in both.iterrows():
        item = {
            "code": code,
            "name": r["SEC_NAME"] if pd.notna(r["SEC_NAME"]) else code,
            "np0": r[b] / 1e8,
            "np1": r[f] / 1e8,
            "delta": r["delta"] / 1e8,
            "g": ratio(r[f], r[b]) - 1,
            "pe": ratio(r["MKT_CAP_ARD"], r[f]),
            "rev": {},
        }
        for w in windows:
            past = r.get(core.np_col(core.FORECAST_YEAR, w))
            item["rev"][w] = (ratio(r[f], past) - 1) if pd.notna(past) and r[f] > 0 else None
        rows.append({k: (clean(v) if k != "rev" else {x: clean(y) for x, y in v.items()})
                     for k, v in item.items()})
    return rows


def industry_record(row, g, windows):
    by, fy, n = core.BASE_YEAR, core.FORECAST_YEAR, core.TOP_N
    rec = {
        "name": row["行业"],
        "n": row["成分股数"],
        "nf": row[f"有{fy}预测家数"],
        "nboth": row["共同样本家数"],
        "cov": row["预测样本市值覆盖率(已知市值口径)"],
        "np0": row[f"{by}E净利润(亿,共同样本)"],
        "np1": row[f"{fy}E净利润(亿,共同样本)"],
        "delta": row[f"{fy}E增量(亿,共同样本)"],
        "g": row[f"{fy}E增速(共同样本)"],
        "gmed": row[f"个股{fy}E增速中位数(基期利润>0)"],
        "pe": row[f"{fy}E PE(市值利润齐全样本)"],
        "inst": row[f"平均机构数({fy}预测样本)"],
        "top3": row[f"前{n}公司{fy}E利润占比"],
        "top3names": row[f"前{n}公司"],
        "turn": row["扭亏贡献占比"],
        "maxco": row["最大增量公司"],
        "maxshare": row["最大增量公司贡献占比"],
        "gex": row["剔除最大增量公司后增速"],
        "flags": [x for x in str(row["提示"]).split("、") if x],
    }
    rec = {k: clean(v) if k != "flags" else v for k, v in rec.items()}
    rec["rev"] = {
        w: {
            "r": clean(row[f"{fy}E修正_{w}(配对样本)"]),
            "amt": clean(row[f"{fy}E修正额_{w}(亿)"]),
            "n": clean(row[f"修正样本家数_{w}"]),
            "up": clean(row[f"上调家数_{w}"]),
            "down": clean(row[f"下调家数_{w}"]),
            "net": clean(row[f"上调净占比_{w}"]),
        }
        for w in windows
    }
    rec["stocks"] = stock_rows(g, windows)
    if "剔除基期偏高公司后增速" in row:   # 港股：基期一次性因素
        rec["oneoff"] = {"names": clean(row["基期偏高公司"]), "share": clean(row["基期偏高利润占比"]),
                         "g": clean(row["剔除基期偏高公司后增速"])}
    return rec


def build_payload(xlsx, title):
    sheets = pd.read_excel(xlsx, sheet_name=None)
    details = sheets["个股明细"].copy()
    details["股票代码"] = details["股票代码"].astype(str)
    details = details.set_index("股票代码")
    windows = configure_core(details)

    summary = sheets.get("行业汇总")
    extra_cols = ["基期偏高公司", "基期偏高利润占比", "剔除基期偏高公司后增速", "提示"]
    hk_extra = (summary.set_index("行业")[[c for c in extra_cols if c in summary.columns]]
                if summary is not None and "剔除基期偏高公司后增速" in summary.columns else None)
    industries = []
    for name, g in details.groupby("行业", sort=False):
        row = core.summarize(g, name)
        if hk_extra is not None and name in hk_extra.index:
            # 港股特有指标与提示（财年错位、币种缺失、基期含一次性）取自 Excel 行业汇总
            row.update({k: v for k, v in hk_extra.loc[name].items() if not (k == "提示" and pd.isna(v))})
        industries.append(industry_record(row, g, windows))

    notes = sheets.get("统计说明")
    note_lines = notes.iloc[:, 0].dropna().astype(str).tolist() if notes is not None else []
    window_dates = {}
    for line in note_lines:
        for w, dt in re.findall(r"([^：；，\s]+?)前=(\d{4}-\d{2}-\d{2})", line):
            window_dates[w] = dt
    date = str(details["请求数据日期"].dropna().iloc[0]) if "请求数据日期" in details else ""
    if not any("提示标签" in line for line in note_lines):
        # 旧版 Excel 的说明不含新指标，按当前口径重新生成
        note_lines = core.note_lines({w: window_dates.get(w, "—") for w in windows}, date)

    overall = overall_stats(details, windows)
    overall = {k: clean(v) if k != "rev" else {x: clean(y) for x, y in v.items()}
               for k, v in overall.items()}
    return {
        "title": title,
        "currency": "港元" if "港股" in title else "",
        "date": date,
        "baseYear": core.BASE_YEAR,
        "forecastYear": core.FORECAST_YEAR,
        "topN": core.TOP_N,
        "windows": windows,
        "windowDates": window_dates,
        "overall": overall,
        "industries": industries,
        "notes": note_lines,
    }


def build(xlsx, out_html, title="申万一级行业"):
    payload = build_payload(xlsx, title)
    data = json.dumps(payload, ensure_ascii=False, allow_nan=False).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8")
    html = html.replace("__TITLE__", f"{title} {payload['forecastYear']}E 盈利预期")
    html = html.replace("/*__DATA__*/null", data)
    Path(out_html).parent.mkdir(parents=True, exist_ok=True)
    Path(out_html).write_text(html, encoding="utf-8")
    return out_html


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="由盈利预期 Excel 生成 HTML 看板")
    ap.add_argument("xlsx")
    ap.add_argument("html", nargs="?")
    ap.add_argument("--title", help="看板标题，默认按文件名判断（港股 / 申万一级）")
    args = ap.parse_args()
    src = Path(args.xlsx)
    dst = Path(args.html) if args.html else src.with_suffix(".html")
    title = args.title or (
        ("港股 · Wind二级行业" if "Wind二级" in src.name else "港股 · 恒生一级行业")
        if src.name.startswith("港股") else "申万一级行业"
    )
    print(f"看板已保存：{build(src, dst, title)}")
