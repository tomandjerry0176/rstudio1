"""指数层面 2027E 盈利预期（自下而上）：读取申万 / 港股脚本输出的个股明细，按指数成分股与权重汇总。

依赖：WindPy（只取指数成分股与权重）、pandas、openpyxl；以及已生成的两份 Excel：
    申万一级_2027盈利预期_<DATE>.xlsx            （sw1_consensus_2027.py）
    港股_恒生一级_2027盈利预期_<DATE>.xlsx        （hk_consensus_2027.py）
用法：
    python index_consensus.py                         # 默认读 OUT_DIR 下与 DATE 同日的两份 Excel
    python index_consensus.py --a A.xlsx --h H.xlsx   # 指定文件

口径：
  指数权重法（主口径）：成分股 i 的盈利收益率 EY = 利润 / 总市值，指数盈利 E = Σ wᵢ·EYᵢ（wᵢ 为指数权重），
    即按指数持有份额加总利润，与 BBG 指数 EPS 的构造一致；增速 = E₂₀₂₇ / E₂₀₂₆ − 1，PE = Σw / Σw·EY₂₀₂₇。
    个股贡献 = wᵢ(EY₂₀₂₇ − EY₂₀₂₆) / E₂₀₂₆（百分点），行业贡献为其加总，合计等于指数增速。
  合计法（对照）：成分股利润直接相加，与行业汇总表口径相同。
  核心增速：剔除基期亏损（扭亏）公司与基期偏高（可能含一次性收益）公司后的权重法增速。
  预期修正：同一批股票、当前市值下，Σw·EY₂₀₂₇（当前）/ Σw·EY₂₀₂₇（对比日）− 1。
"""
import argparse
import json
import re
from pathlib import Path

import numpy as np
import pandas as pd

DATE = "2026-11-03"
BASE_YEAR = 2026
FORECAST_YEAR = 2027
OUT_DIR = Path("/Users/amber/Documents/本地数据库/财报")
TOP_STOCKS = 15
ONE_OFF_DROP = 0.30   # A 股基期偏高：FY2027 预期较 FY2026 下降超过 30%（港股直接用明细中的“基期提示”）
TEMPLATE = Path(__file__).with_name("index_dashboard_template.html")

# code 为 None 表示用明细中的全部股票、按总市值加权；codes 列出候选 Wind 代码，取第一个返回成分股的
INDICES = [
    {"name": "沪深300", "market": "A", "codes": ["000300.SH"]},
    {"name": "中证A500", "market": "A", "codes": ["000510.SH", "000510.CSI"]},
    {"name": "创业板指", "market": "A", "codes": ["399006.SZ"]},
    {"name": "科创50", "market": "A", "codes": ["000688.SH"]},
    {"name": "全部A股（申万）", "market": "A", "codes": None},
    {"name": "恒生指数", "market": "H", "codes": ["HSI.HI"]},
    {"name": "恒生科技", "market": "H", "codes": ["HSTECH.HI"]},
    {"name": "恒生综合", "market": "H", "codes": ["HSCI.HI"]},
]


def np_col(year, window=None):
    return f"NP{year}" if window is None else f"NP{year}_{window}前"


def ratio(num, den):
    if pd.isna(num) or pd.isna(den) or den <= 0:
        return float("nan")
    return num / den


def clean(v):
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (float, np.floating)):
        return None if not np.isfinite(v) else round(float(v), 6)
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


# ---------------------------------------------------------------- 读取个股明细
def load_details(path, market):
    sheets = pd.read_excel(path, sheet_name=None)
    d = sheets["个股明细"].copy()
    d["股票代码"] = d["股票代码"].astype(str).str.strip().str.upper()
    d = d.set_index("股票代码")
    windows = [m.group(1) for c in d.columns if (m := re.match(rf"NP{FORECAST_YEAR}_(.+)前$", c))]
    b, f = np_col(BASE_YEAR), np_col(FORECAST_YEAR)
    if "基期提示" in d.columns:
        flag = d["基期提示"].fillna("").astype(str).ne("")
    else:
        flag = (d[b] > 0) & d[f].notna() & (d[f] < d[b] * (1 - ONE_OFF_DROP))
    d["基期偏高"] = flag
    d["市场"] = market
    dates = {}
    notes = sheets.get("统计说明")
    if notes is not None:
        for line in notes.iloc[:, 0].dropna().astype(str):
            for w, dt in re.findall(r"([^：；，\s]+?)前=(\d{4}-\d{2}-\d{2})", line):
                dates[w] = dt
    return d, windows, dates


# ---------------------------------------------------------------- 成分股与权重
def fetch_members(w, idx):
    """返回 (实际代码, Series[代码 -> 权重(归一化)], 权重方式)。"""
    for code in idx["codes"]:
        err, cons = w.wset("indexconstituent", f"date={DATE};windcode={code}", usedf=True)
        if err != 0 or cons is None or cons.empty:
            continue
        cons = cons.copy()
        cons.columns = cons.columns.astype(str).str.lower()
        if "wind_code" not in cons.columns:
            continue
        cons["wind_code"] = cons["wind_code"].astype(str).str.strip().str.upper()
        cons = cons.drop_duplicates("wind_code").set_index("wind_code")
        wt = pd.to_numeric(cons["i_weight"], errors="coerce") if "i_weight" in cons.columns else None
        if wt is not None and wt.notna().sum() >= 0.9 * len(cons) and wt.sum() > 0:
            return code, (wt.fillna(0) / wt.sum()), "指数权重"
        return code, pd.Series(np.nan, index=cons.index), "总市值"
    raise RuntimeError(f"{idx['name']}：候选代码 {idx['codes']} 都未返回成分股，请在 Wind 中核对指数代码")


# ---------------------------------------------------------------- 计算
def analyze(name, code, market, weights, wmode, d, windows):
    members = weights.index
    have = members.intersection(d.index)
    g = d.loc[have].copy()
    if wmode == "总市值":
        wt = g["MKT_CAP_ARD"].where(g["MKT_CAP_ARD"] > 0)
        weights = (wt / wt.sum()).reindex(members)
    g["权重"] = weights.reindex(have)
    b, f = np_col(BASE_YEAR), np_col(FORECAST_YEAR)
    mv = g["MKT_CAP_ARD"].where(g["MKT_CAP_ARD"] > 0)
    g["EY0"], g["EY1"] = g[b] / mv, g[f] / mv
    s = g.dropna(subset=["EY0", "EY1", "权重"])
    e0, e1 = (s["权重"] * s["EY0"]).sum(), (s["权重"] * s["EY1"]).sum()
    growth = ratio(e1, e0) - 1
    pe_s = g.dropna(subset=["EY1", "权重"])
    core = s[(s[b] > 0) & ~s["基期偏高"]]
    c0, c1 = (core["权重"] * core["EY0"]).sum(), (core["权重"] * core["EY1"]).sum()

    row = {
        "指数": name, "代码": code or "—", "市场": "A股" if market == "A" else "港股", "权重方式": wmode,
        "成分股数": len(members), "明细中找到": len(have), "共同样本家数": len(s),
        "共同样本权重覆盖率": float(s["权重"].sum()) if wmode == "指数权重" else ratio(s["MKT_CAP_ARD"].sum(), g["MKT_CAP_ARD"].sum()),
        f"{FORECAST_YEAR}E增速(指数权重法)": growth,
        f"{FORECAST_YEAR}E增速(合计法)": ratio(s[f].sum(), s[b].sum()) - 1,
        f"{FORECAST_YEAR}E核心增速(剔除扭亏与基期偏高)": ratio(c1, c0) - 1,
        "核心样本权重占比": ratio(core["权重"].sum(), s["权重"].sum()),
        f"{FORECAST_YEAR}E PE(指数权重法)": ratio(pe_s["权重"].sum(), (pe_s["权重"] * pe_s["EY1"]).sum()),
        f"{FORECAST_YEAR}E净利润合计(亿,共同样本)": s[f].sum() / 1e8,
    }
    s = s.copy()
    s["增速贡献"] = s["权重"] * (s["EY1"] - s["EY0"]) / e0 if e0 > 0 else np.nan
    rev = {}
    for win in windows:
        p_col = np_col(FORECAST_YEAR, win)
        g[f"EYp_{win}"] = g[p_col] / mv
        p = g.dropna(subset=["EY1", f"EYp_{win}", "权重"])
        now, past = (p["权重"] * p["EY1"]).sum(), (p["权重"] * p[f"EYp_{win}"]).sum()
        r = ratio(now, past) - 1 if now > 0 else float("nan")
        row[f"{FORECAST_YEAR}E修正_{win}"] = r
        rev[win] = (p, past)

    # 行业贡献
    ind_rows = []
    for ind, gi in s.groupby("行业", sort=False):
        i0, i1 = (gi["权重"] * gi["EY0"]).sum(), (gi["权重"] * gi["EY1"]).sum()
        item = {
            "指数": name, "行业": ind, "家数": len(gi),
            "权重(共同样本内)": ratio(gi["权重"].sum(), s["权重"].sum()),
            f"{BASE_YEAR}E盈利占比": ratio(i0, e0),
            f"{FORECAST_YEAR}E增速(权重法)": ratio(i1, i0) - 1,
            "增速贡献(百分点)": gi["增速贡献"].sum() * 100,
        }
        for win, (p, past) in rev.items():
            pi = p[p["行业"] == ind]
            item[f"修正贡献_{win}(百分点)"] = ((pi["权重"] * (pi["EY1"] - pi[f"EYp_{win}"])).sum() / past * 100
                                           if past > 0 else float("nan"))
        ind_rows.append(item)
    ind_df = pd.DataFrame(ind_rows).sort_values("增速贡献(百分点)", ascending=False)

    # 个股贡献（按增速贡献绝对值）
    top = s.reindex(s["增速贡献"].abs().sort_values(ascending=False).index).head(TOP_STOCKS)
    stock_rows = []
    for c, r in top.iterrows():
        item = {"指数": name, "股票代码": c, "名称": r["SEC_NAME"], "行业": r["行业"], "权重": r["权重"],
                f"{FORECAST_YEAR}E增速": ratio(r[f], r[b]) - 1 if r[b] > 0 else float("nan"),
                "基期亏损": bool(r[b] <= 0), "基期偏高": bool(r["基期偏高"]),
                "增速贡献(百分点)": r["增速贡献"] * 100}
        for win in windows:
            past = r.get(np_col(FORECAST_YEAR, win))
            item[f"修正_{win}"] = ratio(r[f], past) - 1 if pd.notna(past) and r[f] > 0 else float("nan")
        stock_rows.append(item)
    return row, ind_df, pd.DataFrame(stock_rows)


# ---------------------------------------------------------------- 输出
def note_lines(windows, dates, files):
    return [
        f"数据日期：{DATE}；个股预期来自 {files['A'].name} 与 {files['H'].name}，与行业汇总为同一批数据。",
        "指数权重法（主口径）：EY = 利润 / 总市值，指数盈利 = Σ 权重×EY；增速 = 2027E / 2026E − 1；"
        "PE = Σ权重 / Σ权重×EY2027E。按指数持有份额加总利润，与 BBG 指数 EPS 构造一致。",
        "增速贡献（百分点）= 权重×(EY2027E − EY2026E) / 指数 2026E 盈利；行业贡献为个股加总，各行业合计等于指数增速。",
        "合计法：共同样本利润直接相加，与行业汇总表口径相同；与权重法的差异来自指数权重（自由流通、权重上限等）。",
        "核心增速：剔除 2026E 利润不为正（扭亏）和基期偏高（A 股：2027E 较 2026E 下降超过 30%；港股：个股明细“基期提示”）的公司。",
        "预期修正：同一批股票、按当前市值，Σ权重×EY2027E（当前 / 对比日）− 1；修正贡献同理按行业拆分。"
        + ("对比日期：" + "；".join(f"{w}前={dates.get(w, '—')}" for w in windows) + "。" if windows else ""),
        "权重：Wind indexconstituent 的 i_weight；缺少权重的指数与“全部A股（申万）”按总市值加权（等价于合计法）。",
        "共同样本权重覆盖率 = 2026E、2027E 均有预测且有市值的成分股权重之和；未覆盖部分不参与计算，不按零处理。",
        "港股利润已按报表币种折算为港元并日历化（见港股 Excel 统计说明）；A+H 公司的 EY 用全公司利润 / A+H 合计市值。",
    ]


def write_excel(summary, inds, stocks, notes, out_file):
    with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="指数汇总", index=False)
        inds.to_excel(xw, sheet_name="行业贡献", index=False)
        stocks.to_excel(xw, sheet_name="个股贡献", index=False)
        pd.DataFrame({"统计说明": notes}).to_excel(xw, sheet_name="统计说明", index=False)
        for name, frame in (("指数汇总", summary), ("行业贡献", inds), ("个股贡献", stocks)):
            ws = xw.sheets[name]
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            for col, title in enumerate(frame.columns, start=1):
                if any(k in title for k in ("增速", "覆盖率", "占比", "修正_", "权重")) and "百分点" not in title \
                        and title != "权重方式":
                    for cells in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                        cells[0].number_format = "0.00%"
                elif "百分点" in title:
                    for cells in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                        cells[0].number_format = "0.00"


def build_html(summary, inds, stocks, notes, windows, dates, out_html):
    fy = FORECAST_YEAR
    recs = []
    for _, r in summary.iterrows():
        name = r["指数"]
        ii = inds[inds["指数"] == name]
        ss = stocks[stocks["指数"] == name]
        recs.append({
            "name": name, "code": r["代码"], "market": r["市场"], "wmode": r["权重方式"],
            "n": clean(r["成分股数"]), "nboth": clean(r["共同样本家数"]), "cov": clean(r["共同样本权重覆盖率"]),
            "g": clean(r[f"{fy}E增速(指数权重法)"]), "gsum": clean(r[f"{fy}E增速(合计法)"]),
            "gcore": clean(r[f"{fy}E核心增速(剔除扭亏与基期偏高)"]), "pe": clean(r[f"{fy}E PE(指数权重法)"]),
            "rev": {w: clean(r.get(f"{fy}E修正_{w}")) for w in windows},
            "inds": [{"name": x["行业"], "n": clean(x["家数"]), "w": clean(x["权重(共同样本内)"]),
                      "g": clean(x[f"{fy}E增速(权重法)"]), "c": clean(x["增速贡献(百分点)"]),
                      "rc": {w: clean(x.get(f"修正贡献_{w}(百分点)")) for w in windows}} for _, x in ii.iterrows()],
            "stocks": [{"code": x["股票代码"], "name": clean(x["名称"]) or x["股票代码"], "ind": x["行业"],
                        "w": clean(x["权重"]), "g": clean(x[f"{fy}E增速"]), "c": clean(x["增速贡献(百分点)"]),
                        "loss": bool(x["基期亏损"]), "oneoff": bool(x["基期偏高"]),
                        "rev": {w: clean(x.get(f"修正_{w}")) for w in windows}} for _, x in ss.iterrows()],
        })
    payload = {"date": DATE, "baseYear": BASE_YEAR, "forecastYear": fy, "windows": windows,
               "windowDates": dates, "indices": recs, "notes": notes}
    data = json.dumps(payload, ensure_ascii=False, allow_nan=False).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/null", data)
    Path(out_html).write_text(html, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="指数层面 2027E 盈利预期（自下而上）")
    ap.add_argument("--a", help="申万一级 Excel（默认 OUT_DIR/申万一级_2027盈利预期_<DATE>.xlsx）")
    ap.add_argument("--h", help="港股恒生一级 Excel（默认 OUT_DIR/港股_恒生一级_2027盈利预期_<DATE>.xlsx）")
    args = ap.parse_args()
    files = {
        "A": Path(args.a) if args.a else OUT_DIR / f"申万一级_{FORECAST_YEAR}盈利预期_{DATE}.xlsx",
        "H": Path(args.h) if args.h else OUT_DIR / f"港股_恒生一级_{FORECAST_YEAR}盈利预期_{DATE}.xlsx",
    }
    for k, p in files.items():
        if not p.exists():
            raise SystemExit(f"找不到{'A股' if k == 'A' else '港股'}明细：{p}（先运行对应脚本，或用 --a / --h 指定）")
    details, windows, dates = {}, None, {}
    for k, p in files.items():
        details[k], win, dt = load_details(p, k)
        windows = win if windows is None else [w for w in windows if w in win]
        dates.update(dt)

    from WindPy import w
    if w.start().ErrorCode != 0 or not w.isconnected():
        raise RuntimeError("Wind 未连接，请先登录终端")

    rows, inds, stocks = [], [], []
    for idx in INDICES:
        d = details[idx["market"]]
        if idx["codes"] is None:
            code, weights, wmode = None, pd.Series(np.nan, index=d.index), "总市值"
        else:
            code, weights, wmode = fetch_members(w, idx)
        row, ind_df, st_df = analyze(idx["name"], code, idx["market"], weights, wmode, d, windows)
        rows.append(row)
        inds.append(ind_df)
        stocks.append(st_df)
        missing = row["成分股数"] - row["明细中找到"]
        print(f"{idx['name']}（{code or '明细全部'}，{wmode}）：成分股 {row['成分股数']}，共同样本 {row['共同样本家数']}，"
              f"覆盖 {row['共同样本权重覆盖率']:.1%}，2027E 增速 {row[f'{FORECAST_YEAR}E增速(指数权重法)']:+.1%}"
              + (f"；{missing} 只不在明细中" if missing else ""), flush=True)

    summary, inds, stocks = pd.DataFrame(rows), pd.concat(inds), pd.concat(stocks)
    notes = note_lines(windows, dates, files)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / f"指数_{FORECAST_YEAR}盈利预期_{DATE}.xlsx"
    write_excel(summary, inds, stocks, notes, out)
    print(f"已保存：{out}")
    for html in (out.with_suffix(".html"), OUT_DIR / f"指数_{FORECAST_YEAR}盈利预期_latest.html"):
        build_html(summary, inds, stocks, notes, windows, dates, html)
        print(f"看板已保存：{html}")


if __name__ == "__main__":
    main()
