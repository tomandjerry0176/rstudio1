"""生成 BBG Excel 公式版核对模板：在 Bloomberg 终端的 Excel 中打开、刷新，即得指数 2027E 增速的三口径对照。

用法：python make_bbg_template.py --wind 港股_恒生一级_2027盈利预期_2026-11-04.xlsx [--index "HSTECH Index"] [--out 模板.xlsx]
模板内容：
  说明      ：操作步骤与字段说明
  核对      ：B1 = 指数代码；每行一个成分股，BDP 取 BBG 经调整 EPS（BEST_EPS）、GAAP EPS（BEST_EPS_GAAP）的 FY1/FY2、
              财年截止日与股价，并按 Wind 代码查本套 Wind 利润；顶部汇总三种口径的权重法增速与覆盖率
  成分      ：BDS(指数, INDX_MWEIGHT) 自动展开成分股与权重
  Wind明细  ：本套港股 Excel 的个股明细（代码、名称、市值、2026E/2027E 利润、口径），用于对照
公式口径与 index_consensus.py 一致：指数盈利 E = Σ 权重 × EPS / 股价（或 利润 / 市值），增速 = E(FY2) / E(FY1) − 1，
个股贡献 = 权重 × (EY2 − EY1) / E1 × 100（百分点），合计等于增速。
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MAX_ROWS = 120          # 成分股上限（恒生科技 30、恒生指数约 90）
FIRST = 8               # 核对表数据起始行
HEAD_FILL = PatternFill("solid", fgColor="F0EFEC")
BOLD = Font(bold=True)


def build(wind_xlsx, index, out):
    w = pd.read_excel(wind_xlsx, sheet_name="个股明细")
    keep = [c for c in ["股票代码", "SEC_NAME", "MKT_CAP_ARD", "NP2026", "NP2027", "口径", "FY2026来源", "行业"]
            if c in w.columns]
    w = w[keep]
    wb = Workbook()

    # ---------------- 说明
    s = wb.active
    s.title = "说明"
    lines = [
        "BBG 个股一致预期核对模板（Excel 公式版）",
        "",
        "步骤",
        "1. 在装有 Bloomberg Excel 插件的电脑上打开本文件，启用编辑与宏/加载项。",
        "2. Bloomberg 选项卡 → Refresh Workbook（或 Ctrl+Alt+Shift+R），等待所有 #N/A Requesting Data 消失。",
        "3. 「成分」表 A2 的 BDS 会展开成分股与权重；「核对」表逐行取数，顶部 C2:E4 即三种口径的增速。",
        "4. 保存后把文件发回（保存时 Excel 会缓存公式结果，无需粘贴为值）。",
        "",
        "换指数：改「核对」表 B1（如 HSI Index），刷新即可；成分股超过 120 只需向下复制公式行。",
        "",
        "字段（如返回 #N/A Invalid Field，请在终端 FLDS 核对后改「核对」表第 7 行对应的字段名）",
        "BEST_EPS：BEst 标准口径 EPS，中国互联网公司多为经调整（non-GAAP）",
        "BEST_EPS_GAAP：GAAP 口径 EPS",
        "BEST_PERIOD_END_DT：该财年截止日（看 3 月年结公司的 FY1 是否已滚到 2027-03）",
        "覆盖参数：BEST_FPERIOD_OVERRIDE=1FY / 2FY，EQY_FUND_CRNCY=HKD（EPS 换成港元，与股价一致）",
        "",
        "口径（与 index_consensus.py 相同）",
        "指数盈利 E = Σ 权重 × EPS / 股价；增速 = E(FY2) / E(FY1) − 1；个股贡献 = 权重 × (EY2 − EY1) / E1 × 100（百分点）",
        "Wind 列：EY = 本套 2026E、2027E 利润（港元、日历年）/ 总市值，来自「Wind明细」表",
        "三口径都只用 FY1、FY2 均有数的股票，权重覆盖列为其权重之和",
        "",
        "检查项：某只股票 BBG 与 Wind 的 EY 相差数倍时，多为每 ADS / 每股口径或币种问题，请标出发回。",
    ]
    for i, t in enumerate(lines, start=1):
        s.cell(row=i, column=1, value=t)
    s["A1"].font = Font(bold=True, size=13)
    for r in (3, 11, 17):
        s.cell(row=r, column=1).font = BOLD
    s.column_dimensions["A"].width = 110

    # ---------------- 成分
    m = wb.create_sheet("成分")
    m["A1"], m["B1"] = "成分股（BDS 自动展开）", "权重(%)"
    m["A1"].font = m["B1"].font = BOLD
    m["A2"] = '=BDS(核对!$B$1,"INDX_MWEIGHT")'
    m.column_dimensions["A"].width = 22

    # ---------------- Wind明细
    wd = wb.create_sheet("Wind明细")
    for j, c in enumerate(keep, start=1):
        wd.cell(row=1, column=j, value=c).font = BOLD
    for i, row in enumerate(w.itertuples(index=False), start=2):
        for j, v in enumerate(row, start=1):
            wd.cell(row=i, column=j, value=None if pd.isna(v) else v)
    wd_last = len(w) + 1
    wcol = {c: get_column_letter(j) for j, c in enumerate(keep, start=1)}

    def wind_lookup(code_cell, col):
        rng = f"Wind明细!${wcol[col]}$2:${wcol[col]}${wd_last}"
        key = f"Wind明细!$A$2:$A${wd_last}"
        hit = f"INDEX({rng},MATCH({code_cell},{key},0))"
        return f'IFERROR(IF({hit}="","",{hit}),"")'   # 空单元格 INDEX 会返回 0，需转成空

    # ---------------- 核对
    k = wb.create_sheet("核对", 1)
    k["A1"], k["B1"] = "指数代码", index
    k["A1"].font = BOLD
    k["B1"].fill = PatternFill("solid", fgColor="FFF2CC")
    k["D1"] = "← 改这里换指数，再 Refresh Workbook"

    cols = [
        ("成分(BBG)", None), ("BBG代码", None), ("名称", None), ("权重", None), ("Wind代码", None),
        ("股价", "PX_LAST"), ("经调整EPS_FY1", "BEST_EPS"), ("经调整EPS_FY2", "BEST_EPS"),
        ("GAAP_EPS_FY1", "BEST_EPS_GAAP"), ("GAAP_EPS_FY2", "BEST_EPS_GAAP"),
        ("财年截止_FY1", "BEST_PERIOD_END_DT"), ("财年截止_FY2", "BEST_PERIOD_END_DT"),
        ("Wind口径", None), ("Wind_FY2026来源", None),
        ("经调整EY1", None), ("经调整EY2", None), ("GAAP_EY1", None), ("GAAP_EY2", None),
        ("Wind_EY2026", None), ("Wind_EY2027", None),
        ("有效_经调整", None), ("有效_GAAP", None), ("有效_Wind", None),
        ("贡献_经调整(pp)", None), ("贡献_GAAP(pp)", None), ("贡献_Wind(pp)", None), ("贡献差_Wind−经调整(pp)", None),
    ]
    L = {name: get_column_letter(j) for j, (name, _) in enumerate(cols, start=1)}
    for j, (name, fld) in enumerate(cols, start=1):
        c = k.cell(row=FIRST - 1, column=j, value=name)
        c.font, c.fill = BOLD, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if fld:
            k.cell(row=FIRST - 2, column=j, value=fld).font = Font(color="808080", size=9)
    k.cell(row=FIRST - 2, column=1, value="BBG 字段 →").font = Font(color="808080", size=9)

    first, last = FIRST, FIRST + MAX_ROWS - 1
    for r in range(first, last + 1):
        src = r - FIRST + 2                       # 成分表从第 2 行开始展开
        A = lambda n: f"{L[n]}{r}"                # noqa: E731
        k[A("成分(BBG)")] = f'=IF(成分!A{src}="","",成分!A{src})'
        k[A("BBG代码")] = f'=IF({A("成分(BBG)")}="","",{A("成分(BBG)")}&" Equity")'
        k[A("名称")] = f'=IF({A("BBG代码")}="","",BDP({A("BBG代码")},"NAME"))'
        k[A("权重")] = f'=IF({A("BBG代码")}="","",成分!B{src}/SUM(成分!$B$2:$B${MAX_ROWS + 1}))'
        k[A("Wind代码")] = (f'=IF({A("成分(BBG)")}="","",TEXT(VALUE(LEFT({A("成分(BBG)")},'
                            f'FIND(" ",{A("成分(BBG)")})-1)),"0000")&".HK")')
        bdp = lambda fld, fy: (f'=IF({A("BBG代码")}="","",BDP({A("BBG代码")},"{fld}",'  # noqa: E731
                               f'"BEST_FPERIOD_OVERRIDE","{fy}","EQY_FUND_CRNCY","HKD"))')
        k[A("股价")] = f'=IF({A("BBG代码")}="","",BDP({A("BBG代码")},"PX_LAST"))'
        k[A("经调整EPS_FY1")] = bdp("BEST_EPS", "1FY")
        k[A("经调整EPS_FY2")] = bdp("BEST_EPS", "2FY")
        k[A("GAAP_EPS_FY1")] = bdp("BEST_EPS_GAAP", "1FY")
        k[A("GAAP_EPS_FY2")] = bdp("BEST_EPS_GAAP", "2FY")
        k[A("财年截止_FY1")] = bdp("BEST_PERIOD_END_DT", "1FY")
        k[A("财年截止_FY2")] = bdp("BEST_PERIOD_END_DT", "2FY")
        k[A("Wind口径")] = "=" + wind_lookup(A("Wind代码"), "口径") if "口径" in wcol else ""
        if "FY2026来源" in wcol:
            k[A("Wind_FY2026来源")] = "=" + wind_lookup(A("Wind代码"), "FY2026来源")
        ey = lambda eps: f'=IFERROR({A(eps)}/{A("股价")},"")'  # noqa: E731
        k[A("经调整EY1")] = ey("经调整EPS_FY1")
        k[A("经调整EY2")] = ey("经调整EPS_FY2")
        k[A("GAAP_EY1")] = ey("GAAP_EPS_FY1")
        k[A("GAAP_EY2")] = ey("GAAP_EPS_FY2")
        mv = wind_lookup(A("Wind代码"), "MKT_CAP_ARD")
        k[A("Wind_EY2026")] = f'=IFERROR({wind_lookup(A("Wind代码"), "NP2026")}/{mv},"")'
        k[A("Wind_EY2027")] = f'=IFERROR({wind_lookup(A("Wind代码"), "NP2027")}/{mv},"")'
        ok = lambda a, b: (f'=AND(ISNUMBER({A("权重")}),ISNUMBER({A(a)}),ISNUMBER({A(b)}))')  # noqa: E731
        k[A("有效_经调整")] = ok("经调整EY1", "经调整EY2")
        k[A("有效_GAAP")] = ok("GAAP_EY1", "GAAP_EY2")
        k[A("有效_Wind")] = ok("Wind_EY2026", "Wind_EY2027")

    rng = lambda n: f"${L[n]}${first}:${L[n]}${last}"  # noqa: E731

    def e(flag, ey):
        # SUMPRODUCT 逐项相乘时把文本（空行的 ""）当 0；有效标记保证计入的行两期都有数
        return f"SUMPRODUCT(--({rng(flag)}=TRUE),{rng('权重')},{rng(ey)})"

    # 汇总：增速 = Σw·EY2 / Σw·EY1 − 1
    k["A2"], k["B2"], k["C2"], k["D2"], k["E2"] = "口径", "Σw·EY(FY1)", "增速(权重法)", "权重覆盖", "Σw·EY(FY2)"
    for c in ("A2", "B2", "C2", "D2", "E2"):
        k[c].font, k[c].fill = BOLD, HEAD_FILL
    specs = [
        (3, "BBG 经调整 EPS（FY2/FY1，相对财年）", "有效_经调整", "经调整EY1", "经调整EY2"),
        (4, "BBG GAAP EPS（FY2/FY1，相对财年）", "有效_GAAP", "GAAP_EY1", "GAAP_EY2"),
        (5, "Wind 归母净利润（2027E/2026E，日历年）", "有效_Wind", "Wind_EY2026", "Wind_EY2027"),
    ]
    for row, label, flag, ey1, ey2 in specs:
        k[f"A{row}"] = label
        k[f"B{row}"] = "=" + e(flag, ey1)
        k[f"E{row}"] = "=" + e(flag, ey2)
        k[f"C{row}"] = f'=IFERROR(E{row}/B{row}-1,"")'
        k[f"D{row}"] = f'=SUMIF({rng(flag)},TRUE,{rng("权重")})'
        k[f"C{row}"].number_format = "+0.0%;-0.0%"
        k[f"D{row}"].number_format = "0%"
    for r in range(first, last + 1):
        A = lambda n: f"{L[n]}{r}"  # noqa: E731
        k[A("贡献_经调整(pp)")] = (f'=IF({A("有效_经调整")},{A("权重")}*({A("经调整EY2")}-{A("经调整EY1")})/$B$3*100,"")')
        k[A("贡献_GAAP(pp)")] = (f'=IF({A("有效_GAAP")},{A("权重")}*({A("GAAP_EY2")}-{A("GAAP_EY1")})/$B$4*100,"")')
        k[A("贡献_Wind(pp)")] = (f'=IF({A("有效_Wind")},{A("权重")}*({A("Wind_EY2027")}-{A("Wind_EY2026")})/$B$5*100,"")')
        k[A("贡献差_Wind−经调整(pp)")] = f'=IFERROR({A("贡献_Wind(pp)")}-{A("贡献_经调整(pp)")},"")'
        k[A("权重")].number_format = "0.00%"
        for n in ("财年截止_FY1", "财年截止_FY2"):
            k[A(n)].number_format = "yyyy-mm-dd"
        for n in ("贡献_经调整(pp)", "贡献_GAAP(pp)", "贡献_Wind(pp)", "贡献差_Wind−经调整(pp)"):
            k[A(n)].number_format = "0.00"
    k["G2"] = "贡献合计校验（应等于增速×100）"
    k["G2"].font = BOLD
    for row, col in ((3, "贡献_经调整(pp)"), (4, "贡献_GAAP(pp)"), (5, "贡献_Wind(pp)")):
        k[f"G{row}"] = f"=SUM({rng(col)})"
        k[f"G{row}"].number_format = "0.00"
    k.column_dimensions["A"].width = 38
    for j in range(2, len(cols) + 1):
        k.column_dimensions[get_column_letter(j)].width = 13
    k.freeze_panes = f"D{FIRST}"
    wb.save(out)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成 BBG Excel 公式版核对模板")
    ap.add_argument("--wind", required=True)
    ap.add_argument("--index", default="HSTECH Index")
    ap.add_argument("--out")
    a = ap.parse_args()
    out = Path(a.out) if a.out else Path(a.wind).with_name(f"BBG核对模板_{a.index.split()[0]}.xlsx")
    print(f"已生成：{build(a.wind, a.index, out)}")
