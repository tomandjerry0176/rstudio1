"""生成港股（恒生综合全部成分股）BBG 一致预期取数模板（Excel 公式版）。

用法：python make_bbg_hk_template.py --wind 港股_恒生一级_2027盈利预期_<日期>.xlsx [--out 模板.xlsx]
股票清单、数据日期与 1/2/3 个月对比日期都取自 Wind 港股 Excel，保证与 Wind 口径同一批股票、同一组日期。
在 Bloomberg 终端的 Excel 中打开 → Refresh Workbook → 保存发回，再用 hk_bbg_merge.py 合并。

取数（全部 EQY_FUND_CRNCY=HKD，金额单位为百万港元，BBG 默认）：
  BEST_NET_INCOME      1FY / 2FY / 3FY（当前）            BDP
  EQY_FISCAL_YR_END    公司财年截止日（如 03/31），仅作交叉核对；1FY 的截止年份由合并脚本按财年截止月推算
  BEST_NET_INCOME      1FY / 2FY / 3FY（1/2/3 个月前）    BDH，起止日期同一天，Days=A、Fill=P 取当日或之前最近值
"""
import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="F0EFEC")
INPUT = PatternFill("solid", fgColor="FFF2CC")
FIRST = 5          # 数据起始行（第 4 行为表头）


def wind_to_bbg(code):
    m = re.match(r"(\d+)\.HK$", str(code).strip().upper())
    return f"{int(m.group(1))} HK Equity" if m else None


def read_wind(path):
    sheets = pd.read_excel(path, sheet_name=None)
    d = sheets["个股明细"]
    date = str(d["请求数据日期"].dropna().iloc[0])[:10] if "请求数据日期" in d else ""
    windows = {}
    for line in sheets.get("统计说明", pd.DataFrame({"x": []})).iloc[:, 0].dropna().astype(str):
        for w, dt in re.findall(r"([^：；，\s]+?)前=(\d{4}-\d{2}-\d{2})", line):
            windows[w] = dt
    return d, date, windows


def build(wind_xlsx, out):
    d, date, windows = read_wind(wind_xlsx)
    if not windows:
        raise SystemExit("Wind Excel 的统计说明里没有找到对比日期（形如 1个月前=2026-09-29）")
    wb = Workbook()

    s = wb.active
    s.title = "说明"
    for i, t in enumerate([
        "港股 BBG 一致预期取数模板（恒生综合全部成分股，Excel 公式版）",
        "",
        "1. 在装有 Bloomberg Excel 插件的电脑上打开，Bloomberg 选项卡 → Refresh Workbook，等待所有 Requesting Data 消失。",
        "   581 只 × 14 个公式，约需几分钟；如提示数据量上限，可分批选中区域刷新（Refresh Selection）。",
        "2. 保存后发回（Excel 会缓存公式结果，无需粘贴为值），用 hk_bbg_merge.py 合并到港股流程。",
        "",
        "字段（#N/A Invalid Field 时在终端 FLDS 核对，改「BBG取数」第 3 行的字段名与对应公式）：",
        "BEST_NET_INCOME：BEst 一致预期净利润（BBG 标准口径，部分公司为经调整），单位百万，已用 EQY_FUND_CRNCY=HKD 换成港元",
        "EQY_FISCAL_YR_END：公司财年截止日，仅作核对；1FY 对应哪一年由合并脚本按财年截止月 + 3 个月报出期推算",
        "BDH 历史值：起止日为「参数」表中的对比日期（与 Wind 修正对比日期一致）；Days=A、Fill=P 取当日或之前最近值，",
        "  Dts=H 只返回数值。若单元格显示的是日期而非数值，说明 Dts=H 未生效，请告知。",
        "",
        "检查：#N/A N.A. = BBG 无该财年预测，属正常；整列 #N/A 多为字段名或覆盖参数问题。",
    ], start=1):
        s.cell(row=i, column=1, value=t)
    s["A1"].font = Font(bold=True, size=13)
    s.column_dimensions["A"].width = 120

    p = wb.create_sheet("参数")
    p["A1"], p["B1"] = "数据日期", pd.Timestamp(date).to_pydatetime()
    rows = list(windows.items())
    for i, (w, dt) in enumerate(rows, start=2):
        p.cell(row=i, column=1, value=f"{w}前")
        p.cell(row=i, column=2, value=pd.Timestamp(dt).to_pydatetime())
    for r in range(1, len(rows) + 2):
        p.cell(row=r, column=1).font = BOLD
        p.cell(row=r, column=2).number_format = "yyyy-mm-dd"
        p.cell(row=r, column=2).fill = INPUT
    p["D1"] = "日期来自 Wind 港股 Excel；如改动，请同步改 Wind 对比日期，否则修正口径不可比"
    p.column_dimensions["A"].width = 14
    p.column_dimensions["B"].width = 14

    k = wb.create_sheet("BBG取数", 1)
    k["A1"] = f"恒生综合成分股 {len(d)} 只 · 数据日期 {date} · 金额：百万港元"
    k["A1"].font = BOLD
    cols = [("Wind代码", None, None), ("BBG代码", None, None), ("名称(Wind)", None, None), ("名称(BBG)", "NAME", None)]
    for fy in ("1FY", "2FY", "3FY"):
        cols.append((f"NI_{fy}", "BEST_NET_INCOME", fy))
    cols.append(("FYE(BBG)", "EQY_FISCAL_YR_END", None))
    for i, (w, _) in enumerate(rows):
        for fy in ("1FY", "2FY", "3FY"):
            cols.append((f"NI_{fy}_{w}前", "BEST_NET_INCOME", (fy, i + 2)))
    for j, (name, fld, _) in enumerate(cols, start=1):
        c = k.cell(row=4, column=j, value=name)
        c.font, c.fill = BOLD, HEAD
        if fld:
            k.cell(row=3, column=j, value=fld).font = Font(color="808080", size=9)
        if isinstance(_, tuple):
            k.cell(row=2, column=j, value=f"=参数!B{_[1]}").number_format = "yyyy-mm-dd"
    k.cell(row=2, column=1, value="BDH 日期 →").font = Font(color="808080", size=9)
    k.cell(row=3, column=1, value="BBG 字段 →").font = Font(color="808080", size=9)

    for r, row in enumerate(d.itertuples(index=False), start=FIRST):
        code = getattr(row, "股票代码")
        k.cell(row=r, column=1, value=code)
        k.cell(row=r, column=2, value=wind_to_bbg(code))
        k.cell(row=r, column=3, value=getattr(row, "SEC_NAME"))
        t = f"$B{r}"
        for j, (name, fld, opt) in enumerate(cols, start=1):
            if j <= 3:
                continue
            if fld in ("NAME", "EQY_FISCAL_YR_END"):
                f = f'=BDP({t},"{fld}")'
            elif isinstance(opt, tuple):
                fy, prow = opt
                dc = f"{get_column_letter(j)}$2"
                f = (f'=BDH({t},"{fld}",{dc},{dc},"BEST_FPERIOD_OVERRIDE","{fy}",'
                     f'"EQY_FUND_CRNCY","HKD","Days=A","Fill=P","Dts=H")')
            else:
                f = f'=BDP({t},"{fld}","BEST_FPERIOD_OVERRIDE","{opt}","EQY_FUND_CRNCY","HKD")'
            k.cell(row=r, column=j, value=f)
    k.freeze_panes = f"D{FIRST}"
    for j in range(1, len(cols) + 1):
        k.column_dimensions[get_column_letter(j)].width = 14
    wb.save(out)
    return out, len(d), windows


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成港股 BBG 一致预期取数模板")
    ap.add_argument("--wind", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()
    out = Path(a.out) if a.out else Path(a.wind).with_name("BBG取数模板_港股恒生综合.xlsx")
    path, n, win = build(a.wind, out)
    print(f"已生成：{path}（{n} 只，对比日期 {win}）")
