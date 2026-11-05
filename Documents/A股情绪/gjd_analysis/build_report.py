"""把分析写进研报模板：python3 build_report.py
模板：../研报模板.docx    输出：../国家队ETF资金跟踪_<最新日期>.docx
正文文字在本文件 TEXT 部分，数据更新后需要同步修改文字；阶段表与图表自动生成。"""
import os, re, shutil, zipfile, tempfile, html
from PIL import Image
from stages import S, d

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, '..', '研报模板.docx')
FIG = os.path.join(HERE, 'report_figs')
OUT = os.path.join(HERE, '..', f'国家队ETF资金跟踪_{d.index[-1]:%Y%m%d}.docx')
REPORT_DATE = '5 Nov 2026'

# ============================ TEXT ============================
TITLE = '国家队重回净买入：本轮宽基ETF回补约3,300亿元'

SUMMARY = [
    ('本轮国家队已净买入约3,320亿元，规模居2023年以来第二。',
     '自9月28日以来，我们识别出15个国家队交易日，扣除跟风资金后累计净买入约3,320亿元，在2023年8月以来的8轮集中买入中仅次于2024年6–9月的3,578亿元，略高于2024年1–2月救市期的3,303亿元。'),
    ('节奏明显加快，11月3日单日买入约610亿元。',
     '11月2–4日三个交易日国家队净买入约1,263亿元，为本轮最强的一周；11月3日单日约610亿元，为2025年4月8日（1,048亿元）以来最大的单日买入。'),
    ('11月4日跟风资金明显增多。',
     '当日份额净流入389亿元，但流入占超额成交的比例降至0.65；中证500/1000小规模产品集中流入约64亿元，创业板ETF成交放大而净流入有限，科创ETF净流出。按成交结构拆分，当日跟风资金约90亿元，国家队约300亿元。'),
    ('品种由大盘向成长扩散。',
     '沪深300 ETF仍占本轮国家队买入的52%，但创业板占比升至17%，明显高于2024–2025年历次托底时的4–6%，延续了7月以来的特征。'),
    ('从仓位看，回补仍处早期。',
     '2026年1–9月国家队在样本ETF上累计净卖出约1.42万亿元，7月初低点以来累计回补约4,700亿元，回补比例约三分之一。'),
]
RISK = '国家队资金识别基于统计规则，可能与实际持仓变动存在偏差；跟风资金拆分依赖参数假设；样本仅覆盖23只主要宽基ETF；政策节奏与市场环境变化。'

INTRO = ('10月中旬以来，A股宽基ETF成交额与份额同步放大，市场对“国家队”是否重新入场的讨论升温。'
         '本文从ETF成交额与ETF份额两条线索交叉印证，识别2023年8月以来国家队资金的进出，复盘历次集中买卖阶段，并对本轮的规模、节奏、结构与跟风资金做定量刻画。')

# 每节：(小标题, [段落...], [图表组...])；段落为 (粗体导语, 正文)；图表组见 build 部分
SECTIONS = [
    ('方法：成交额与份额两条线索交叉印证', [
        ('两条线索。', '其一是成交额：以前120个交易日成交额的20%分位作为“常态成交额”，当日成交额超出常态的部分记为超额成交。'
         '其二是份额：当日ETF场内份额变化乘以当日均价，即申购赎回带来的资金净流入。样本为23只主要宽基ETF，覆盖上证50、沪深300、中证500、中证1000、创业板、科创50等指数。'),
        ('判别逻辑。', '国家队以大额单向申购或赎回为主，成交放大的部分几乎全部转化为份额变化；散户换手放大时买卖双向，净申赎有限。'
         '据此，我们将同时满足“|净流入|＞0.3×常态成交额”与“|净流入|/超额成交≥0.4”的交易日认定为国家队交易日。'
         '相较固定的±50亿元阈值，该规则随市场成交规模自适应，既能捕捉2024年夏季单日30–50亿元的缩量买入，也能过滤高成交期的小额噪音。'),
        ('交叉印证结果。', '在识别出的255个国家队交易日上，净流入与超额成交的相关系数达0.97，二者比值中位数为0.89（图表2）。'
         '对照之下，2024年9月24日至10月24日超额成交约9,466亿元，份额净流入仅约1,017亿元，比值中位数0.34；其中10月9日前净流入1,804亿元，10月10日后转为净流出787亿元，呈现典型的散户追涨后兑现特征，我们将该时段剔除。'
         '参数在合理区间内调整（阈值0.25–0.35、比值0.3–0.4），各阶段金额变化均在±5%以内。'),
    ], [('full', 'fig1_cum.png', '图表1 国家队宽基ETF累计净流入（亿元，2023年8月1日起）'),
        ('pair', ('fig2_scatter.png', '图表2 超额成交与份额净流入：国家队交易日 vs 散户主导期'),
                 ('fig3_stages.png', '图表3 历次集中买入阶段的净买入规模'))]),

    ('复盘：2023年以来的8轮集中买入与3轮集中卖出', [
        ('救市型：2024年1–2月。', '小盘股流动性危机期间，成交额放大至常态的2–3倍，19个交易日中有17天出现国家队买入，2月5–8日连续四天单日买入225–485亿元，累计约3,303亿元。'),
        ('缩量托底型：2024年6–9月。', '平均成交额仅185亿元，却在64个交易日内累计买入约3,578亿元，是样本中规模最大的一轮。单日金额不大（日均约100亿元），但持续时间长、几乎没有卖出，流入/超额成交高达0.90。'),
        ('急救型：2025年4月。', '关税冲击后，4月7–8日两天买入约1,584亿元，4月8日单日1,048亿元为样本最高值，之后小额补仓即结束，强度最高、持续最短。'),
        ('降温型卖出：2026年1–7月。', '1月15日至2月2日连续13个交易日净卖出约9,652亿元，日均742亿元；4–7月又分两轮卖出约5,500亿元。'
         '累计曲线由2025年5月末的约1.0万亿元高点降至7月初的−6,334亿元，说明卖出部分包含样本起点之前已持有的仓位。'),
    ], [('table', '图表4 2023年以来国家队集中买卖阶段一览')]),

    ('本轮：规模、节奏、结构与跟风资金', [
        ('规模居前。', '本轮自9月28日起共23个交易日，其中15个为国家队交易日，扣除跟风资金后净买入约3,320亿元，规模在历次买入中排第二；国家队交易日的单日均额约221亿元，高于2024年1–2月（194亿元）和2024年夏季（97亿元），低于2025年4月的急救期（318亿元）。'),
        ('分批放量，逐波加强。', '本轮可分为三波：10月15–16日约528亿元，10月21–26日约1,000亿元，11月2–4日约1,263亿元（图表5）。'
         '每一波的力度都超过前一波，节奏与2024年初的持续托底相似，而非2025年4月那种一两天的急救。'),
        ('品种向成长扩散。', '沪深300 ETF仍占本轮国家队买入的52%，与历次一致，可以确认是同一类资金；但创业板占比17%，明显高于2024年1–2月、2024年夏季和2025年4月的4–6%，延续了7月那一轮（18%）的特征（图表6）。'),
        ('11月4日跟风资金明显增多。', '我们将超额成交拆为国家队与跟风两部分：纯国家队交易的流入/超额成交约0.9（2024年夏季与2025年4月实测），跟风资金的净流入/超额成交取本轮放量但单向性不足日子的实测值0.34。'
         '据此，11月4日389亿元净流入中，跟风资金约92亿元，国家队约298亿元；若跟风比例取0.21，则跟风约45亿元、国家队约345亿元。'
         '分品种看也有印证：中证500/1000的几只小规模产品当日合计流入约64亿元（此前国家队交易日几乎为零），创业板ETF超额成交152亿元但净流入仅49亿元，科创ETF净流出；而国家队最常用的四只沪深300 ETF流入由11月3日的337亿元降至130亿元。'
         '按同一口径，本轮累计扣除跟风资金约466亿元，约占份额净流入的12%。'),
    ], [('pair', ('fig4_current.png', '图表5 本轮日度净流入拆分：国家队 vs 跟风及其他资金（亿元）'),
                 ('fig5_mix.png', '图表6 主要买入阶段的品种结构（%）'))]),

    ('结论与跟踪指标', [
        ('国家队已由“降温卖出”转向“托底买入”。', '本轮净买入约3,320亿元，力度居2023年以来第二，且呈逐波加强之势；但与2026年1–9月约1.42万亿元的净卖出相比，回补比例仍仅约三分之一。'),
        ('跟风资金占比上升是一把双刃剑。', '一方面说明托底行为正在带动增量资金入场；另一方面，若后续国家队节奏放缓而跟风资金主导，ETF资金流的单向性将下降，需警惕类似2024年10月“追涨—兑现”的反复。'),
        ('跟踪指标。', '一是四只沪深300核心ETF（510300、510310、510330、159919）的份额变化，这是判断国家队是否持续买入最直接的信号；'
         '二是全市场宽基ETF的“流入/超额成交”，若持续低于0.4，意味着成交主导力量正由国家队转向散户。'),
    ], []),
]

EN_SUMMARY = [
    'The "national team" has turned back into a net buyer of A-share broad-based ETFs. Cross-checking ETF turnover against creation/redemption flows, we estimate it has bought about RMB332bn since 28 Sep 2026 (net of follow-on retail flows), the second-largest buying wave since Aug 2023, behind only Jun–Sep 2024 (RMB358bn).',
    'Buying has accelerated: about RMB126bn on 2–4 Nov, including about RMB61bn on 3 Nov, the largest single day since 8 Apr 2025. Follow-on retail money rose noticeably on 4 Nov; we estimate roughly RMB9bn of that day\'s RMB39bn inflow came from retail investors.',
    'CSI 300 ETFs still account for 52% of national-team buying, but the ChiNext share has risen to 17%, versus 4–6% in earlier rescue waves. Relative to roughly RMB1.42tn of net selling in Jan–Sep 2026, the replenishment is still only about one-third complete.',
]
EN_RISK = 'Risk factors: national-team flows are identified by statistical rules and may deviate from actual holdings changes; the split of follow-on flows depends on parameter assumptions; the sample covers 23 major broad-based ETFs only; changes in policy pace and market conditions.'

# ============================ BUILD ============================
FONT = '<w:rFonts w:eastAsia="楷体_GB2312" w:cs="Calibri" w:hint="eastAsia"/>'
def esc(t): return html.escape(t, quote=False)
def run(t, bold=False, color='000000', sz=20):
    return (f'<w:r><w:rPr>{FONT}{"<w:b/>" if bold else ""}<w:bCs/><w:snapToGrid/><w:color w:val="{color}"/>'
            f'<w:kern w:val="2"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
def para(runs, after=100, jc='both', keep=False, ind=None):
    return (f'<w:p><w:pPr>{"<w:keepNext/>" if keep else ""}<w:spacing w:after="{after}" w:line="276" w:lineRule="auto"/>'
            f'{ind or ""}<w:jc w:val="{jc}"/></w:pPr>{"".join(runs)}</w:p>')
def heading(t):  # 沿用模板“一、小标题”的自动编号（numId 13）
    return (f'<w:p><w:pPr><w:pStyle w:val="a8"/><w:keepNext/><w:widowControl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="13"/></w:numPr>'
            f'<w:spacing w:before="240" w:after="120" w:line="278" w:lineRule="auto"/></w:pPr>{run(t, True, "23558C", 22)}</w:p>')

EMU = 914400; rels = []; pic_id = [5000]
def drawing(fn, width_in):
    im = Image.open(os.path.join(FIG, fn)); w = int(width_in * EMU); h = int(w * im.height / im.width)
    rid = f'rId{900 + len(rels)}'; rels.append((rid, fn)); pic_id[0] += 1; pid = pic_id[0]
    return (f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w}" cy="{h}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{pid}" name="图片 {pid}"/><wp:cNvGraphicFramePr>'
            f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="{pid}" name="{fn}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

NOB = ('<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
       '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
       '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>')
def cap(t):  # 图表标题：模板样式 8.5pt 粗体
    return f'<w:p><w:pPr><w:keepNext/><w:spacing w:after="60"/></w:pPr>{run(t, True, "000000", 17)}</w:p>'
def src(t='Source: Wind, HTI'):
    return f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{run(t, False, "000000", 10)}</w:p>'
def cell(w, inner): return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{inner}</w:tc>'
def ctable(cols, rows):
    grid = ''.join(f'<w:gridCol w:w="{c}"/>' for c in cols)
    body = ''.join('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(cell(c, x) for c, x in zip(cols, r)) + '</w:tr>' for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="a9"/><w:tblW w:w="{sum(cols)}" w:type="dxa"/>{NOB}'
            f'<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0"/></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>')
GAP = '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'
def fig_full(fn, title):
    return ctable([10698], [[cap(title)], [f'<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr>{drawing(fn, 7.2)}</w:p>'], [src()]]) + GAP
def fig_pair(a, b):
    return ctable([5349, 5349], [[cap(a[1]), cap(b[1])],
                                [f'<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr>{drawing(a[0], 3.55)}</w:p>',
                                 f'<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr>{drawing(b[0], 3.55)}</w:p>'],
                                [src(), src()]]) + GAP

TYPE = {'③': '救市型', '⑤': '缩量托底型', '⑥': '急救型', 'Ⓐ': '降温型卖出', 'Ⓑ': '分批卖出', 'Ⓒ': '分批卖出', '⑦': '托底', '①': '温和买入', '②': '年末托底', '④': '补仓'}
def stage_table(title):
    cols = [700, 2050, 1150, 1300, 1000, 1100, 1150, 1100, 1148]
    hdr = ['编号', '区间', '净额（亿元）', '国家队交易日/交易日', '日均（亿元）', '最大单日（亿元）', '流入/超额成交', '平均成交额（亿元）', '类型']
    def tc(w, t, head=False, bold=False, last=False):
        shd = '<w:shd w:val="clear" w:color="auto" w:fill="23558C"/>' if head else ''
        bdr = (f'<w:tcBorders><w:bottom w:val="single" w:sz="{8 if (head or last) else 4}" w:space="0" w:color="{"23558C" if (head or last) else "D9D9D9"}"/></w:tcBorders>')
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{bdr}{shd}<w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>{run(t, head or bold, "FFFFFF" if head else "000000", 15)}</w:p></w:tc>')
    rows = ['<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' + ''.join(tc(w, h, head=True) for w, h in zip(cols, hdr)) + '</w:tr>']
    for i, rw in S.iterrows():
        cur = bool(rw.本次)
        vals = [rw.编号, f'{rw.start:%Y/%-m/%-d}–{rw.end:%Y/%-m/%-d}', f'{rw.GJD净额:+,}', f'{rw.信号日}/{rw.交易日}', f'{rw.日均:,}',
                f'{rw.最大单日:,}', f'{abs(rw.流入_超额):.2f}', f'{rw.均成交额:,}',
                ('本轮（已扣跟风%d亿元）' % rw.跟风扣除) if cur else TYPE.get(rw.编号, '')]
        rows.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(tc(w, v, bold=cur, last=(i == len(S) - 1)) for w, v in zip(cols, vals)) + '</w:tr>')
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(cols)}" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="23558C"/></w:tblBorders>'
           f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0"/></w:tblPr>'
           f'<w:tblGrid>{"".join(f"<w:gridCol w:w=\"{c}\"/>" for c in cols)}</w:tblGrid>{"".join(rows)}</w:tbl>')
    note = run('注：买入≥500亿元、卖出≥1,000亿元的阶段；流入/超额成交为阶段内份额净流入与超额成交之比的绝对值。', False, '000000', 10)
    return cap(title) + tbl + f'<w:p><w:pPr><w:spacing w:before="40" w:after="0"/></w:pPr>{note}</w:p>' + src() + GAP

def lead_para(lead, body):
    return para([run(lead, True), run(body)])

body = para([run(INTRO)], after=160)
for h, paras, figs in SECTIONS:
    body += heading(h)
    body += ''.join(lead_para(a, b) for a, b in paras)
    for f in figs:
        if f[0] == 'full': body += fig_full(f[1], f[2])
        elif f[0] == 'pair': body += fig_pair(f[1], f[2])
        elif f[0] == 'table': body += stage_table(f[1])

# ---------- 写入模板 ----------
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(TPL) as z: z.extractall(tmp)
p = os.path.join(tmp, 'word', 'document.xml'); s = open(p, encoding='utf-8').read()

def sub_text(old, new, count=1):
    global s
    assert old in s, old
    s = s.replace(f'<w:t>{old}</w:t>', f'<w:t>{esc(new)}</w:t>', count)
sub_text('13 Jul 2025', REPORT_DATE)
sub_text('标题', TITLE)

def para_bounds(idx):
    a = s.rfind('<w:p ', 0, idx); b = s.find('</w:p>', idx) + 6
    return a, b
# 摘要框：替换“正文”段，删除其后的空段，改写风险提示
i = s.index('<w:t>正文</w:t>'); a, b = para_bounds(i)
summary = ''.join(para([run(l, True, '23558C'), run(t, False, '262626')], after=100) for l, t in SUMMARY)
nxt_a, nxt_b = para_bounds(b + 10)
s = s[:a] + summary + s[nxt_b:]
s = s.replace('<w:t>全球流动性环境变化、海外利率波动、国际资本流动以及个股解禁后的实际减持行为</w:t>', f'<w:t>{esc(RISK)}</w:t>', 1)

# 正文：从“引言”段到正文风险提示段之前，整体替换
i0 = s.index('<w:t>引言</w:t>'); a0, _ = para_bounds(i0)
i1 = s.index('<w:t>全球流动性环境变化、海外利率波动、</w:t>'); a1, b1 = para_bounds(i1)
risk_end = para([run('风险提示：', True, '23558C'), run(RISK, True, '23558C')])
s = s[:a0] + body + risk_end + s[b1:]

# 风险提示后的手动分页符段 → 改为 APPENDIX 1 段前分页（避免正文刚好排满时多出一页空白）
i = s.index('<w:br w:type="page"/>'); a, b = para_bounds(i); s = s[:a] + s[b:]
i = s.index('<w:t>APPENDIX 1</w:t>'); a, _ = para_bounds(i)
j = s.index('<w:pStyle w:val="aa"/>', a) + len('<w:pStyle w:val="aa"/>'); s = s[:j] + '<w:pageBreakBefore/>' + s[j:]   # 须紧跟 pStyle（schema 顺序）

# 英文摘要
i = s.index('Risk factors: changes in global liquidity'); a, b = para_bounds(i)
tpl_par = s[a:b]
def en_par(t): return re.sub(r'<w:t>.*?</w:t>', f'<w:t xml:space="preserve">{esc(t)}</w:t>', tpl_par, count=1)
s = s[:a] + ''.join(en_par(t) for t in EN_SUMMARY + [EN_RISK]) + s[b:]

# 去掉模板里不再使用的示例图片引用（rId10/rId11 的 emf 仍保留在包内无害）
open(p, 'w', encoding='utf-8').write(s)

# 图片与关系
media = os.path.join(tmp, 'word', 'media')
rp = os.path.join(tmp, 'word', '_rels', 'document.xml.rels'); rs = open(rp, encoding='utf-8').read()
add = ''
for rid, fn in rels:
    shutil.copy(os.path.join(FIG, fn), os.path.join(media, 'gjd_' + fn))
    add += f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/gjd_{fn}"/>'
rs = rs.replace('</Relationships>', add + '</Relationships>'); open(rp, 'w', encoding='utf-8').write(rs)

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(tmp):
        for f in files:
            full = os.path.join(root, f); arc = os.path.relpath(full, tmp)
            z.write(full, arc)
shutil.rmtree(tmp)
print('written', os.path.abspath(OUT))
