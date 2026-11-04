"""Build v2 of the Korea household balance sheet report by editing the original docx in place.
Usage: python3 -I build.py <src.docx> <out.docx> <img_margin.png> <img_other.png>
"""
import sys, copy, re
import docx
from docx.shared import Inches
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import parse_xml

SRC, OUT, IMG_MARGIN, IMG_OTHER = sys.argv[1:5]
d = docx.Document(SRC)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def find_p(prefix):
    for p in d.paragraphs:
        if p.text.startswith(prefix):
            return p
    raise KeyError(prefix)

# ---------- element factories (mirror original XML formatting) ----------
def body(text, lead=None):
    runs = ""
    if lead:
        runs += f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{esc(lead)}</w:t></w:r>'
    runs += f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    return parse_xml(f'<w:p {W}>{runs}</w:p>')

def note(text):
    return parse_xml(f'<w:p {W}><w:pPr><w:spacing w:line="276" w:lineRule="auto" w:after="140"/></w:pPr>'
                     f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def heading(text, level):
    return parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r><w:t>{esc(text)}</w:t></w:r></w:p>')

def caption(text):
    return parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="Caption"/><w:spacing w:after="160"/></w:pPr>'
                     f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def empty():
    return parse_xml(f'<w:p {W}/>')

def picture(path):
    p = d.add_paragraph()
    p.add_run().add_picture(path, width=Inches(6.93))
    el = p._p
    el.getparent().remove(el)
    ppr = parse_xml(f'<w:pPr {W}><w:keepNext/><w:spacing w:after="40"/></w:pPr>')
    el.insert(0, ppr)
    return el

BORDER = ('<w:tcBorders><w:top w:val="single" w:sz="4" w:color="D9D9D9"/><w:left w:val="single" w:sz="4" w:color="D9D9D9"/>'
          '<w:bottom w:val="single" w:sz="4" w:color="D9D9D9"/><w:right w:val="single" w:sz="4" w:color="D9D9D9"/></w:tcBorders>'
          '<w:tcMar><w:top w:w="70" w:type="dxa"/><w:bottom w:w="70" w:type="dxa"/><w:left w:w="85" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tcMar>')

def cell(text, width, fill, header, align):
    rpr = '<w:b/><w:color w:val="FFFFFF"/>' if header else '<w:b w:val="0"/><w:color w:val="000000"/>'
    return (f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{width}"/><w:vAlign w:val="center"/>{BORDER}<w:shd w:fill="{fill}"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/><w:jc w:val="{align}"/></w:pPr>'
            f'<w:r><w:rPr>{rpr}<w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p></w:tc>')

def row_xml(vals, widths, idx, header=False, aligns=None):
    fill = "1B4E70" if header else ("FFFFFF" if idx % 2 == 0 else "F0F4F7")
    trpr = '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' if header else '<w:trPr><w:cantSplit/></w:trPr>'
    cells = "".join(cell(v, w, fill, header, (aligns[i] if aligns else ("left" if i == 0 else "center")))
                    for i, (v, w) in enumerate(zip(vals, widths)))
    return f'<w:tr>{trpr}{cells}</w:tr>'

TOTAL = 10037
def table(headers, rows, widths=None, aligns=None):
    n = len(headers)
    if widths is None:
        first = 3557 if n <= 4 else 2400
        rest = (TOTAL - first) // (n - 1)
        widths = [first] + [rest] * (n - 1)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    trs = row_xml(headers, widths, 0, True, aligns) + "".join(row_xml(r, widths, i, False, aligns) for i, r in enumerate(rows))
    return parse_xml(f'<w:tbl {W}><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/>'
                     f'<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>'
                     f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>')

def append_rows(tbl, rows):
    """Append rows to an existing docx table, copying its widths and alternating shading."""
    t = tbl._tbl
    widths = [int(g.get(qn("w:w"))) for g in t.find(qn("w:tblGrid"))]
    existing = len(t.findall(qn("w:tr"))) - 1
    for i, r in enumerate(rows):
        t.append(parse_xml(f'<w:tbl {W}>{row_xml(r, widths, existing + i)}</w:tbl>').find(qn("w:tr")))

def source(num, title, url, desc):
    rid = d.part.relate_to(url, RT.HYPERLINK, is_external=True)
    return parse_xml(f'<w:p {W}><w:pPr><w:spacing w:after="120" w:line="259" w:lineRule="auto"/></w:pPr>'
                     f'<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">[{num}] </w:t></w:r>'
                     f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:color w:val="145B8C"/></w:rPr><w:t>{esc(title)}</w:t></w:r></w:hyperlink>'
                     f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:br/><w:t xml:space="preserve">{esc(desc)}</w:t></w:r></w:p>')

def insert_after(anchor, elements):
    for el in elements:
        anchor.addnext(el)
        anchor = el
    return anchor

def insert_before(anchor, elements):
    for el in elements:
        anchor.addprevious(el)

def set_text(p, lead, text):
    """Replace a paragraph's runs with bold lead + plain text, keeping pPr."""
    for r in list(p._p):
        if r.tag != qn("w:pPr"):
            p._p.remove(r)
    new = body(text, lead)
    for r in list(new):
        p._p.append(r)

# ---------- 0. title / subtitle ----------
sub = d.paragraphs[1]
sub.runs[0].text = "韩国专题更新  2026年11月4日｜增补：股市调整后的散户去杠杆"
d.core_properties.subject = "截至2026年11月4日可得资料的韩国专题更新，增补股市调整与散户去杠杆"

# ---------- 1. core judgments ----------
p_total = find_p("总量缓冲仍在")
insert_before(p_total._p, [
    body("KOSPI自6月19日盘中高点9,385.59点回落，7月、10月两度触发熔断，11月3日收于5,648.79点，仍较高点低39.8%。"
         "信用融资余额从6月24日峰值38.63万亿韩元降至8月4日27.40万亿，随8月反弹回升至10月12日约34.29万亿，"
         "10月30日再降至29.07万亿。7月券商日均强制平仓金额创纪录，杠杆在第一轮出清后很快重建，第二轮仍在进行。[16][17][23][24][26][27][28]",
         "股市去杠杆已走两轮，尚未出清。"),
    body("信用融资峰值仅约为家庭及非营利机构金融负债的1.5%，股市对居民资产负债表的主要冲击在资产端："
         "一季度末居民持有权益证券及基金1,849.1万亿韩元，市值每降10%约对应185万亿韩元账面损失（本稿推算）。"
         "股票相关借款已明显退潮：全金融机构其他贷款5—7月合计增加11.9万亿，8—9月回吐7.0万亿；"
         "但违约指标对股价的反应通常滞后一至两个季度，10月第二轮下跌的影响尚未进入已公布的逾期数据。[4][29][39]",
         "冲击主要在资产端，负债端外溢有限但存在时滞。"),
])
set_text(p_total, "总量缓冲仍在。",
         "2025年底家庭及非营利机构净资产为14,200.5万亿韩元，同比增长9.0%；2026年二季度金融资产与金融负债之比升至2.78倍。"
         "二季度读数截至6月底，未反映7月后的股价下跌；资产市值改善也不能直接推断借款人的可用现金同步增加。[1][2]")
p_new = find_p("新增贷款放缓")
set_text(p_new, "新增贷款放缓，周转压力仍需跟踪。",
         "9月全金融机构家庭贷款净减少2.2万亿韩元，卡贷余额已较5月高点下降2.1%；但现金预借8月末升至2024年末以来最高，"
         "结算型循环信用9月回升至7.052万亿韩元，10月5大银行信用贷款也重新增加。更新后的关注点是新增借款放缓之后的还款与周转能力。[4][5][14][34]")
p_youth = find_p("青年负债需要分层分析")
set_text(p_youth, "青年负债需要分层分析。",
         "29岁及以下、30至39岁家庭的债务资产比均约30.3%，高于全部家庭的16.8%；20—30多岁信用融资余额6月末达4.19万亿韩元，"
         "为2024年末的2.4倍，与其卡贷余额同一量级。20多岁、30多岁的卡贷余额下降与逾期率上升同时发生；"
         "住房贷款借款人、股票杠杆借款人与低信用借款人的风险不同，应分别观察。[8][9][11][35]")

# ---------- 2. key data table + note ----------
append_rows(d.tables[0], [
    ["KOSPI收盘（较盘中高点）", "5,648.79点（-39.8%）", "2026年11月3日"],
    ["证券信用融资余额（较峰值）", "29.07万亿韩元（-24.8%）", "2026年10月30日"],
    ["投资者预托金", "88.68万亿韩元", "2026年10月22日"],
    ["20—30多岁信用融资余额", "4.19万亿韩元", "2026年6月末"],
])
p_scope = find_p("本稿使用截至2026年11月4日可得资料")
p_scope.runs[0].text = (p_scope.runs[0].text +
    "股市与证券信用数据为金融投资协会、韩国交易所口径的媒体转引读数；本次检索环境无法直接打开协会FreeSIS原表，正式引用前建议逐日核对。")

# ---------- 3. section 1 tail paragraph ----------
p_q2 = find_p("这些资产读数截至6月底")
set_text(p_q2, None,
         "这些资产读数截至6月底。7月后股市两轮大幅调整（见第2节），二季度财富缓冲不能直接外推到11月；个别家庭也无法调用其他家庭的净资产偿债。[10]")

# ---------- 4. renumber old sections 2..6 -> 3..7 and figures ----------
for p in d.paragraphs:
    if p.style.name == "Heading 1":
        m = re.match(r"^(\d) (.*)$", p.text)
        if m and int(m.group(1)) >= 2:
            p.runs[0].text = f"{int(m.group(1)) + 1} {m.group(2)}"
            for r in p.runs[1:]:
                r.text = ""
fig_map = {"图3": "图5", "图2": "图4"}
for p in d.paragraphs:
    if p.style.name == "Caption":
        for old, new in fig_map.items():
            if p.text.startswith(old):
                p.runs[0].text = p.runs[0].text.replace(old, new, 1)
                break

# ---------- 5. new section 2 ----------
h_loans = find_p("3 新增家庭贷款与信用卡周转需求")
sec2 = [
    heading("2 股市调整与散户去杠杆", 1),
    body("2026年上半年KOSPI大幅上涨，5月6日收于7,384.56点，年内涨幅71.4%，居主要市场首位；6月19日盘中创9,385.59点历史高点，当日收于9,052.42点。"
         "此后指数两度急跌：7月月内四次触发熔断，28、29日KOSPI与KOSDAQ史上首次连续两日熔断，7月30日收于5,593.56点的阶段低点；"
         "8月反弹后，10月在AI投资放缓及半导体竞争担忧下再跌，15日单日下跌7.57%并触发熔断，22日收于5,533.77点。"
         "11月3日外资与机构合计净卖出约2.0万亿韩元、个人净买入1.66万亿，指数收于5,648.79点。[15][16][17][18][19][20][21][22]"),
    empty(),
    note("—"),  # placeholder replaced by table below
    note("较高点变化按6月19日盘中高点9,385.59点计算。来源：[15]—[22]，媒体转引韩国交易所数据；10月盘中最低约5,306点，具体日期未能核对，不列入表内。"),
    heading("信用融资：两轮去杠杆之间迅速再加杠杆", 2),
    body("证券信用融资余额6月24日创38.63万亿韩元历史峰值。7月急跌后，8月4日降至27.40万亿，较峰值减少11.2万亿（-29.1%）。"
         "随着指数反弹，余额三周内回升20.8%至8月26日33.10万亿，10月12日约34.29万亿，几乎回到峰值的九成。"
         "10月下跌后余额再降，30日为29.07万亿，较峰值低24.8%；11月2日小幅回升至约29.2万亿，其中KOSPI市场22.58万亿、KOSDAQ市场6.64万亿。[23][24][25][26][27]"),
    picture(IMG_MARGIN),
    caption("图2  证券信用融资余额。来源：[23]—[27]，金融投资协会数据经媒体转引；7月31日约28.9万亿为四舍五入值。"),
    body("余额降幅不等于强制平仓规模：余额同时反映自愿还款、卖出、新增融资与强平。现有证据只支持两点判断：第一，强平在两轮下跌中明显放大了卖压；"
         "第二，8月反弹后散户迅速重建杠杆，10月初的融资余额已接近峰值，第一轮去杠杆并未改变杠杆偏好。"),
    heading("强制平仓、账户损失与现金缓冲", 2),
    body("10家券商信用交易及担保贷款的强制平仓，7月日均438.7亿韩元、涉及2,258个账户，约为上年同月的13倍，远高于2022年9月日均178.9亿的前高；"
         "1月、5月、6月日均分别约29亿、82亿、204亿韩元。金融投资协会另以委托交易未收款口径统计，9月21日至10月21日强平合计5,337亿韩元，"
         "10月14日单日893亿；未收款余额升至1.39万亿韩元，较7月27日的1.08万亿明显增加。两种口径覆盖范围不同，不能相加。[26][28][29]"),
    body("金融当局对两大券商约460万个人账户的抽样显示，3月1—9日的下跌中，使用信用融资账户的平均收益率为-19.0%，未使用账户为-8.2%；"
         "投入不足1,000万韩元的信用账户为-20.7%，同组未使用账户为-7.5%。样本早于6—7月高点，但显示同一波下跌中杠杆账户损失约为非杠杆账户的2倍以上，"
         "小额杠杆账户受损更重。网络流传的“120万账户追保、32—36万账户强平”未见金融当局原始公告，本稿不采用。[29]"),
    body("投资者预托金在6月创历史高点后回落：7月27日为109.17万亿韩元，9月末104.66万亿，10月22日降至88.68万亿，接近年内低点（1月13日88.05万亿），"
         "4个多月减少约51万亿。预托金下降既可能是资金被用于补仓或抄底，也可能是资金离场，单凭总额无法区分。[29][30][31]"),
    heading("银行与非银信贷：股票相关借款先涨后退", 2),
    body("全金融机构“其他贷款”（含信用贷款）2025年全年合计净减少13.9万亿韩元，单月最高仅+1.3万亿；2026年5—7月却连续增加5.3、3.8、2.8万亿，三个月合计11.9万亿。"
         "韩银将5月银行家庭贷款增加6.9万亿、为2024年8月以来最大，归因于家庭月季节性需求叠加个人大规模股票投资；金融当局亦认为5月第二金融圈卡贷与信用贷款增量有相当部分流入股市。"
         "8—9月其他贷款合计减少7.0万亿，回吐5—7月增量的约59%；韩银称9月银行其他贷款减少2.3万亿主要受个人股票投资减少影响。[4][32][33][39]"),
    picture(IMG_OTHER),
    caption("图3  全金融机构其他贷款月度增减。来源：[39]用户提供的金融委月报数据（2025年1月至2026年2月），2026年3—9月采用[4]的9月发布批次。"),
    body("10月出现反向信号：5大银行个人信用贷款增加0.67万亿韩元至108.90万亿，为三个月来首次增加，增幅远小于5—6月的每月2万亿以上；"
         "媒体认为部分与股价下跌后利用透支账户抄底有关。全金融机构10月数据预计11月中旬公布，是判断“再加杠杆”是否从证券账户扩散到银行信贷的关键。[34]"),
    heading("对居民资产负债表的含义", 2),
    body("信用融资峰值38.63万亿韩元约为二季度末家庭及非营利机构金融负债（2,495.0万亿）的1.5%、家庭信用（2,019.8万亿）的1.9%。"
         "即使大部分损失由借款人承担，也不足以单独改变居民部门总量偿债能力；证券信用以账户内股票为抵押、由券商逐日盯市，坏账首先表现为借款人净资产缩水，而非金融机构损失。[2][3]",
         "规模：股票杠杆是放大器，不是主要债务存量。"),
    body("一季度末居民持有权益证券及投资基金1,849.1万亿韩元，占金融资产28.8%。二季度金融资产增加513.1万亿，交易流量约89万亿（净金融投资60.7万亿加负债增量约28.2万亿的近似），"
         "约八成来自估值上升；同季居民净投资权益及基金101.6万亿、存款净减少23.3万亿，说明居民在高点前后从存款转向股票。以一季度末持有额为基数，"
         "权益类资产市值每下降10%约对应185万亿韩元账面损失。该口径含非上市股权和基金，实际弹性低于KOSPI跌幅；三季度资金循环将给出实际估值变化。以上为推算，不是官方统计。[2][29]",
         "资产端：财富效应逆转是主要渠道。"),
    body("2025年调查显示，净资产最高20%家庭持有约一半金融资产；2025年末韩国预托结算院登记的股票所有者为1,455.8万名（含法人、外国人）。"
         "股票财富集中于缓冲较厚的家庭，大部分估值损失落在偿债能力较强的人群；但杠杆账户与小额账户的损失率更高，更可能转向信用贷款、透支账户和卡贷周转。"
         "因此需要观察的是杠杆借款人的二次融资和逾期，而不是总量损失本身。[8][29]",
         "分布：损失集中在高净值家庭，压力集中在杠杆借款人。"),
]
insert_before(h_loans._p, sec2)

# replace placeholder note with KOSPI table
ph = [p for p in d.paragraphs if p.text == "—"][0]
kospi_tbl = table(["时点", "KOSPI", "较高点", "说明"], [
    ["2026年5月6日", "7,384.56", "-21.3%", "年内涨幅71.4%"],
    ["2026年6月19日（盘中）", "9,385.59", "—", "历史高点，收盘9,052.42"],
    ["2026年7月30日", "5,593.56", "-40.4%", "7月阶段收盘低点"],
    ["2026年10月12日", "6,384.24", "-32.0%", "第二轮下跌前"],
    ["2026年10月22日", "5,533.77", "-41.0%", "单日-3.52%"],
    ["2026年11月3日", "5,648.79", "-39.8%", "单日-3.59%"],
], widths=[2900, 2000, 1700, 3437])
ph._p.addprevious(kospi_tbl)
ph._p.getparent().remove(ph._p)

# ---------- 6. card subsection: short-term card credit table ----------
p_card_after = find_p("9月卡贷余额较5月峰值")
card_tbl = table(["产品余额", "2024年末", "2025年末", "2026年4月", "2026年5月", "2026年8月", "2026年9月"], [
    ["卡贷", "42.387", "42.329", "42.983", "43.253", "42.685", "42.340"],
    ["现金预借", "6.948", "6.173", "6.197", "6.504", "6.904", "6.617"],
    ["结算型循环信用", "7.067", "6.720", "6.707", "6.800", "6.999", "7.052"],
], widths=[2237, 1300, 1300, 1300, 1300, 1300, 1300])
insert_after(p_card_after._p, [
    body("拉长时间看，现金预借余额4—8月增加11.4%，8月末6.904万亿韩元为2024年末以来最高，恰在7月暴跌与追保期间；9月回落4.2%。"
         "结算型循环信用5—9月增加3.7%，9月已回到2024年末水平附近。“卡贷收缩、短期现金周转上升”的组合与急需短期流动性的借款人行为一致，"
         "但不能直接证明资金用于补仓。[5][14][39]"),
    empty(), card_tbl,
    note("单位：万亿韩元，月末余额。2024年末至2026年5月来自[39]用户提供的9家卡机构数据（部分为媒体转述的四舍五入值）；8月为韩联社转引，9月为协会表2。范围接近但非完全同一批次。"),
])

# ---------- 7. youth stock leverage (in section 6) ----------
p_y_after = find_p("余额单位：万亿韩元。8家专营卡公司")
youth_tbl = table(["年龄", "2024年末", "2026年6月末", "倍数"], [
    ["20至29岁", "0.194", "0.477", "2.5倍"],
    ["30至39岁", "1.567", "3.718", "2.4倍"],
    ["20至39岁合计", "1.761", "4.194", "2.4倍"],
])
insert_after(p_y_after._p, [
    heading("青年的股票杠杆", 2),
    body("国会议员获得的监管资料显示，6月末20—30多岁信用融资余额4.19万亿韩元，为2024年末1.76万亿的2.4倍；其中30多岁3.72万亿，占88.6%。"
         "另一份10家券商资料显示，20多岁信用融资4月第2周较上年同期增长124%，增速居各年龄组之首。按同期市场余额约38万亿粗算，2030约占一成；"
         "两份资料的券商范围可能不同，该比重仅供参考。[35][36]"),
    empty(), youth_tbl,
    note("单位：万亿韩元。来源：[35]，金融当局提供给国会议员资料经媒体转引。7月后分年龄信用融资未取得，无法判断青年在两轮强平中的去杠杆程度。"),
    body("6月末2030信用融资余额（4.19万亿）与7月末2030卡贷余额（4.96万亿）属同一量级。二者口径不同，但说明青年的股票杠杆已不是边缘现象；"
         "若强平损失转为卡贷、现金预借周转，会与本节所述青年卡贷逾期上升相互强化。"),
])

# ---------- 8. youth tail: personal rehabilitation ----------
p_tail_after = find_p("来源：[12]，金融委提供给国会议员资料")
insert_after(p_tail_after._p, [
    body("司法债务处置也在增加：1—7月20—30多岁个人回生申请29,393人，占全部申请30.4%，同期获确定债务调整的2030为31,590人；"
         "全国个人回生申请全年预计首次超过16万件。国会质询中，釜山2030个人回生申请4年增至2.5倍，投资损失是否计入可减免债务成为争议。"
         "投资损失在申请原因中的占比没有统计，不能把回生增加全部归因于股市。[37][38]"),
])

# ---------- 9. implications ----------
p_impl = find_p("区分两类偿债压力")
insert_before(p_impl._p, [
    body("信用融资只占家庭负债的很小比例，但它通过强平放大指数跌幅，并在反弹时迅速重建。11月初余额约29万亿韩元，仍高于8月初27.4万亿的低点；"
         "若指数再跌而余额不降，强平风险会再次集中释放。",
         "股市杠杆尚未出清。"),
])
set_text(p_impl, "区分三类偿债压力。",
         "青年住房贷款借款人的主要风险，是债务规模与资产积累、稳定收入及还款期限之间的匹配；股票杠杆借款人的风险，是强平损失转化为信用贷款和卡贷周转；"
         "低信用青年的主要风险，是高成本周转与逾期持续。不能用其中一类借款人的数据代表全部青年。")
p_cond = find_p("风险升级需要多项证据同时出现")
set_text(p_cond, "风险升级需要多项证据同时出现。",
         "若循环信用和现金预借继续上升、青年逾期继续扩大，且银行和非银行非房贷逾期同时走高，应上调对消费和弱信用群体的压力判断。"
         "若再叠加股价二次探底、房价走弱和融资条件收紧，才有更充分依据讨论向更广泛金融风险传导。以上为条件判断，不是已发生事实。")
watch_tbl = table(["指标", "发布时间", "需要警惕的信号"], [
    ["信用融资余额与强平金额", "每日，金融投资协会", "指数下跌而余额不降；单日强平再超800亿韩元"],
    ["10月全金融机构家庭贷款", "11月中旬，金融委", "其他贷款重新转正并超过+2万亿韩元"],
    ["10月卡贷、现金预借、循环信用", "11月下旬，信贷金融协会", "现金预借与循环信用同步上升"],
    ["三季度资金循环", "预计2027年1月前后，韩银", "权益资产估值损失规模与存款回流"],
    ["分年龄卡贷逾期与个人回生", "不定期，国会资料", "2030逾期率继续上升"],
], widths=[3200, 2900, 3937])
insert_after(p_cond._p, [heading("后续观察清单", 2), empty(), watch_tbl,
                         note("发布时间按以往惯例推断，以机构实际公布为准。")])

# ---------- 10. gaps table ----------
gaps = [t for t in d.tables if t.rows[0].cells[0].text.strip() == "项目"][0]
append_rows(gaps, [
    ["强平账户数与亏损分布", "无官方常设统计；不采用网传“120万追保、32—36万强平”"],
    ["7月后分年龄信用融资", "未取得，不外推青年去杠杆程度"],
    ["三季度居民权益估值损失", "仅给敏感性，不估算实际值"],
    ["证券信用逐日原表", "本次无法直接访问协会FreeSIS，采用媒体转引读数"],
])

# ---------- 11. sources ----------
SRC_LIST = [
    (15, "时代日报 2026年5月6日KOSPI收盘", "https://www.sidae.com/article/2026050614484038518", "2026年5月6日；KOSPI收于7,384.56点，年内涨幅71.35%。"),
    (16, "Money Today 2026年6月19日市况", "https://www.mt.co.kr/stock/2026/06/19/2026061913361016576", "2026年6月19日；盘中9,385.59点历史高点，收盘9,052.42点。"),
    (17, "韩国日报 7月KOSPI跌幅", "https://www.hankookilbo.com/news/article/A2026080211080002205", "2026年8月2日；7月下跌22%以上，月内四次熔断。"),
    (18, "NEWNEEK 两天连续熔断", "https://newneek.co/@newneek/article/43685", "2026年7月；7月28—29日KOSPI与KOSDAQ连续两日熔断。"),
    (19, "韩国经济 2026年10月13日市况", "https://www.hankyung.com/article/2026101390676", "2026年10月13日；收于6,035.35点，前日6,384.24点；7月30日收盘低点5,593.56点。"),
    (20, "Seoul Economic Daily 10月15日急跌", "https://en.sedaily.com/finance/2026/10/15/kospi-plunges-as-forced-sales-hit-leveraged-retail-investors", "2026年10月15日；单日-7.57%并触发熔断，强平增加。"),
    (21, "Money Today 2026年10月22日收盘", "https://www.mt.co.kr/stock/2026/10/22/2026102215280317031", "2026年10月22日；KOSPI收于5,533.77点。"),
    (22, "NewsPim 2026年11月3日收盘", "https://www.newspim.com/news/view/20261103001168", "2026年11月3日；KOSPI收于5,648.79点，外资、机构净卖出，个人净买入。"),
    (23, "SBS Biz 信用融资降至33万亿元台", "https://v.daum.net/v/20260720164509555", "2026年7月20日；6月24日峰值38.6328万亿，7月16日33.3624万亿。金融投资协会数据。"),
    (24, "The Fact 8月信用融资回升", "https://news.tf.co.kr/read/ogmeta/2359110.htm", "2026年8月；8月4日27.4038万亿，8月26日33.1023万亿。"),
    (25, "Bloomingbit 7月信用融资", "https://en.bloomingbit.io/feed/news/117666", "2026年8月；7月31日约28.9万亿韩元。"),
    (26, "亚洲经济 一个月强平5,300亿", "https://www.ajunews.com/view/20261023145948356", "2026年10月23日；9月21日至10月21日强平5,336.8亿韩元，未收款1.3926万亿；10月12日、16日融资余额。"),
    (27, "PRESS9 信用融资较高点减少9万亿", "https://www.press9.kr/news/articleView.html?idxno=85536", "2026年11月；10月22日30.3633万亿，10月30日29.0664万亿，11月2日分市场余额。"),
    (28, "Metro Seoul 7月强平日均439亿", "https://www.metroseoul.co.kr/article/20260902500413", "2026年9月2日；10家券商信用交易及担保贷款强平日均金额与账户数。"),
    (29, "用户整理：韩国散户持仓与金融资产分布汇总", "https://www.yna.co.kr/amp/view/AKR20260321049300002", "2026年7月28日工作簿；韩银一季度资金循环、KSD所有者、金融投资协会7月27日资金动向、两大券商460万账户样本（链接为韩联社样本报道）。"),
    (30, "Herald Economy 投资者预托金年内低位", "https://biz.heraldcorp.com/article/10914747", "2026年10月；10月22日88.6841万亿韩元，1月13日88.0489万亿。"),
    (31, "Supple 预托金4个月减少51万亿", "https://supple.kr/news/cmvlo0nyr000b4i1ti6pg8uzb", "2026年10月；9月末104.6644万亿，10月以来流出逾15万亿。"),
    (32, "Money Today 5月家庭贷款6.9万亿", "https://www.mt.co.kr/economy/2026/06/11/2026061110490079047", "2026年6月11日；韩银5月金融市场动向，个人股票投资推升其他贷款。"),
    (33, "SBS Biz 9月银行家庭贷款", "https://biz.sbs.co.kr/amp/article/20000340091", "2026年10月15日；韩银9月金融市场动向，其他贷款减少与个人股票投资减少。"),
    (34, "Global Economic 10月5大银行家庭贷款", "https://www.g-enews.com/article/Finance/2026/11/202611021533035850cd0bfacc1c_1", "2026年11月2日；家庭贷款+2.7717万亿，个人信用贷款+0.6687万亿至108.899万亿。"),
    (35, "The Public 2030信用融资4.19万亿", "https://www.thepublic.kr/news/articleView.html?idxno=315571", "2026年；金融当局提供给金炫廷议员的6月末分年龄信用融资资料。"),
    (36, "Asia Today 20多岁信用融资增长124%", "https://www.asiatoday.co.kr/kn/view.php?key=20260429010009482", "2026年4月29日；10家券商资料，提供给姜珉局议员。"),
    (37, "Financial News 2030个人回生近3万人", "https://www.fnnews.com/news/202609240553393115", "2026年9月24日；1—7月2030个人回生申请与债务调整人数。"),
    (38, "News1 釜山2030个人回生", "https://www.news1.kr/local/busan-gyeongnam/6325846", "2026年；釜山2030个人回生4年增至2.5倍，投资损失减免争议。"),
    (39, "用户整理：韩国居民信贷月度数据包", "https://www.fsc.go.kr/no010101/87297", "2026年7月14日工作簿；金融委2025年1月至2026年6月家庭贷款月度增减、9家卡机构2024年12月至2026年5月余额。"),
]
p_scope_note = find_p("口径提示：")
insert_before(p_scope_note._p, [source(*s) for s in SRC_LIST])
p_scope_note.runs[-1].text = (p_scope_note.runs[-1].text +
    "证券信用融资、投资者预托金与银行信贷是不同统计体系，不能互相替换或相加；[15]—[38]多为媒体转引，[29][39]为用户提供的整理工作簿。")

d.save(OUT)
print("saved", OUT)
