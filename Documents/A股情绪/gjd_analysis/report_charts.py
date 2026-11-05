"""研报用图（藏青/灰配色，与研报模板一致）。输出到 report_figs/ 目录。"""
from stages import *
import os
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.dates as mdates

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_figs')
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({'font.sans-serif': ['WenQuanYi Zen Hei'], 'axes.unicode_minus': False,
                     'font.size': 7.5, 'axes.linewidth': 0.6})
NAVY = '#23558C'; GRAY = '#8C8C8C'; LGRAY = '#C8C8C8'; ACC = '#C55A11'; INK = '#262626'; GRID = '#E6E6E6'
DPI = 220

def style(ax):
    ax.grid(axis='y', color=GRID, lw=0.6); ax.set_axisbelow(True)
    for s in ['top', 'right']: ax.spines[s].set_visible(False)
    for s in ['left', 'bottom']: ax.spines[s].set_color('#9A9A9A')
    ax.tick_params(colors=INK, labelsize=7, length=2)

def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=DPI, bbox_inches='tight', facecolor='white')
    plt.close(fig)

buys = S[S.GJD净额 > 0]

# ---- 图1 累计净流入（整页宽） ----
fig, ax = plt.subplots(figsize=(7.4, 2.9)); style(ax)
for _, rw in S.iterrows():
    ax.axvspan(pd.Timestamp(rw.start), pd.Timestamp(rw.end) + pd.Timedelta(days=1),
               color=NAVY if rw.GJD净额 > 0 else GRAY, alpha=0.13 if rw.GJD净额 > 0 else 0.22, lw=0)
ax.axvspan(pd.Timestamp('2024-09-24'), pd.Timestamp('2024-10-24'), facecolor='none', edgecolor=LGRAY, hatch='////', lw=0)
ax.plot(d.index, d.cum, color=NAVY, lw=1.4)
ax.axhline(0, color='#7F7F7F', lw=0.6)
top = d.cum.max() * 1.42; bot = d.cum.min() * 1.6
ax.set_ylim(bot, top)
prev = {True: None, False: None}; nsell = 0
for _, rw in S.iterrows():
    s0, e = pd.Timestamp(rw.start), pd.Timestamp(rw.end); xm = s0 + (e - s0) / 2
    up = rw.GJD净额 > 0
    v = f'{rw.GJD净额:+,}'.replace('-', '−') if abs(rw.GJD净额) >= 1000 else ''
    lab = rw.编号 + ('本轮' if rw.本次 else '')
    near = prev[up] is not None and (xm - prev[up]).days < 75; prev[up] = xm
    if up:
        ax.text(xm, top * 0.97, lab + ('\n' + v if v else ''), ha='center', va='top', fontsize=6.8,
                color=ACC if rw.本次 else NAVY)
    else:   # 卖出阶段：单行标注，交替错开高度
        ax.text(xm, bot * 0.97 + abs(bot) * 0.11 * (nsell % 2), (lab + ' ' + v).strip(), ha='center', va='bottom',
                fontsize=6.8, color='#595959'); nsell += 1
ax.text(pd.Timestamp('2024-10-08'), bot * 0.45, '散户主导期\n(剔除)', ha='center', fontsize=6.5, color='#595959')
ax.annotate(f'{d.index[-1]:%-m/%-d}: {d.cum.iloc[-1]:,.0f}', (d.index[-1], d.cum.iloc[-1]), xytext=(-62, 22),
            textcoords='offset points', fontsize=7, color=INK, arrowprops=dict(arrowstyle='-', color=GRAY, lw=0.6))
ax.set_ylabel('亿元', fontsize=7, color=INK)
ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7]))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))
save(fig, 'fig1_cum.png')

# ---- 图2 超额成交 vs 份额流入：国家队交易日 vs 2024/9/24–10/24 ----
fig, ax = plt.subplots(figsize=(3.5, 2.3)); style(ax)
o = d[(~d.gjd_day) & (~d.retail) & d.flow.notna()]
g = d[d.gjd_day]; rt = d[d.retail]
ax.scatter(o.ex, o.flow, s=4, color=LGRAY, lw=0, label='其他交易日')
ax.scatter(g.ex, g.flow, s=6, color=NAVY, lw=0, label='国家队交易日')
ax.scatter(rt.ex, rt.flow, s=12, color=ACC, lw=0, marker='D', label='2024/9/24–10/24')
xx = np.array([0, d.ex.max()])
ax.plot(xx, 0.9 * xx, color=NAVY, lw=0.6, ls='--'); ax.plot(xx, -0.9 * xx, color=NAVY, lw=0.6, ls='--')
ax.text(1500, 1500 * 0.9 + 80, '流入=0.9×超额成交', fontsize=6, color=NAVY, ha='right')
ax.axhline(0, color='#7F7F7F', lw=0.5)
ax.set_xlabel('超额成交额（亿元）', fontsize=7); ax.set_ylabel('份额×均价净流入（亿元）', fontsize=7)
ax.legend(frameon=False, fontsize=6.3, loc='lower left', handletextpad=0.2)
save(fig, 'fig2_scatter.png')

# ---- 图3 历次买入阶段规模 ----
fig, ax = plt.subplots(figsize=(3.5, 2.3)); style(ax); ax.grid(axis='y', visible=False); ax.grid(axis='x', color=GRID, lw=0.6)
lab = [f'{rw.编号} {pd.Timestamp(rw.start):%y/%-m}–{pd.Timestamp(rw.end):%-m/%-d}' for _, rw in buys.iterrows()]
cols = [ACC if b else NAVY for b in buys.本次]
yy = np.arange(len(buys))[::-1]
ax.barh(yy, buys.GJD净额, color=cols, height=0.62)
for y, v, n in zip(yy, buys.GJD净额, buys.信号日):
    ax.text(v + 40, y, f'{v:,}（{n}天）', va='center', fontsize=6.3, color=INK)
ax.set_yticks(yy); ax.set_yticklabels(lab, fontsize=6.5)
ax.set_xlim(0, buys.GJD净额.max() * 1.32); ax.set_xlabel('净买入（亿元）；括号内为国家队交易日数', fontsize=6.8)
save(fig, 'fig3_stages.png')

# ---- 图4 本轮日度：国家队 vs 跟风/其他 ----
cur = S[S.本次]
s0 = pd.Timestamp(cur.start.iloc[0]) if len(cur) else d.index[-25]
x = d.loc[s0 - pd.Timedelta(days=3):]
other = (x.flow.fillna(0) - x.gjd).clip(lower=0)
fig, ax = plt.subplots(figsize=(3.5, 2.3)); style(ax)
ix = np.arange(len(x))
ax.bar(ix, x.gjd.clip(lower=0), color=NAVY, width=0.7, label='国家队')
ax.bar(ix, other, bottom=x.gjd.clip(lower=0), color=LGRAY, width=0.7, label='跟风及其他资金')
neg = x.flow.clip(upper=0); ax.bar(ix, neg, color=LGRAY, width=0.7)
last = -9; lifted = False
for i, (gv, ov) in enumerate(zip(x.gjd, other)):
    if gv >= 290:
        off = 40 if i - last == 1 and not lifted else 10
        lifted = off == 40
        ax.text(i, gv + ov + off, f'{gv:.0f}', ha='center', fontsize=5.8, color=NAVY); last = i
ax.set_xticks(ix[::3]); ax.set_xticklabels([f'{t:%-m/%-d}' for t in x.index[::3]], fontsize=6.5)
ax.axhline(0, color='#7F7F7F', lw=0.5); ax.set_ylabel('亿元', fontsize=7)
ax.set_ylim(top=max(x.flow.max(), 0) * 1.18)
ax.legend(frameon=False, fontsize=6.5, loc='upper left')
save(fig, 'fig4_current.png')

# ---- 图5 买入阶段的品种结构 ----
grp = {'沪深300': ['510300.SH', '510310.SH', '510330.SH', '159919.SZ'],
       '上证50等': ['510050.SH', '510100.SH', '560050.SH', '510180.SH', '510230.SH'],
       '中证500': ['510500.SH', '159922.SZ', '512500.SH'],
       '中证1000': ['512100.SH', '159845.SZ', '560010.SH', '159629.SZ'],
       '创业板': ['159915.SZ', '159952.SZ', '159977.SZ'],
       '科创及其他': ['588050.SH', '588080.SH', '159901.SZ', '515800.SH']}
comp = {}
for _, rw in buys.iterrows():
    if rw.GJD净额 < 1000: continue
    m = d.loc[str(rw.start):str(rw.end)]; m = m.index[m.gjd_day]
    t = r.loc[m].sum(); comp[rw.编号 + ('本轮' if rw.本次 else '')] = {k: max(t[v].sum(), 0) for k, v in grp.items()}
C = pd.DataFrame(comp).T; C = C.div(C.sum(1), axis=0) * 100
fig, ax = plt.subplots(figsize=(3.5, 2.3)); style(ax); ax.grid(axis='y', visible=False)
pal = ['#23558C', '#6F93BD', '#A9BFD9', '#8C8C8C', ACC, '#D9D9D9']
left = np.zeros(len(C)); yy = np.arange(len(C))[::-1]
for (k, col) in zip(C.columns, pal):
    ax.barh(yy, C[k], left=left, color=col, height=0.6, label=k, edgecolor='white', lw=0.6)
    for y, l, v in zip(yy, left, C[k]):
        if v >= 8: ax.text(l + v / 2, y, f'{v:.0f}', ha='center', va='center', fontsize=6, color='white' if col in (NAVY, '#6F93BD', '#8C8C8C', ACC) else INK)
    left += C[k].values
ax.set_yticks(yy); ax.set_yticklabels(C.index, fontsize=7); ax.set_xlim(0, 100); ax.set_xlabel('占比（%）', fontsize=7)
ax.legend(frameon=False, fontsize=6, ncol=3, loc='upper center', bbox_to_anchor=(0.5, -0.22))
save(fig, 'fig5_mix.png')
C.round(1).to_csv(os.path.join(OUT, 'mix.csv'), encoding='utf-8-sig')
print('ok', OUT)
