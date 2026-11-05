from stages import *
import os
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.dates as mdates
plt.rcParams['font.sans-serif']=['WenQuanYi Zen Hei']; plt.rcParams['axes.unicode_minus']=False
RED='#d63c3c'; BLUE='#2a78d6'; INK='#0b0b0b'; INK2='#52514e'; GRID='#e4e3df'; GRAY='#c9c8c3'
fig,(ax,bx)=plt.subplots(2,1,figsize=(15,9.5),sharex=True,gridspec_kw=dict(height_ratios=[2.2,1],hspace=0.08))
fig.patch.set_facecolor('#fcfcfb')
for a in (ax,bx):
    a.set_facecolor('#fcfcfb'); a.grid(axis='y',color=GRID,lw=0.8); a.set_axisbelow(True)
    for s in ['top','right']: a.spines[s].set_visible(False)
    for s in ['left','bottom']: a.spines[s].set_color(GRAY)
    a.tick_params(colors=INK2,labelsize=10)
    a.axvspan(pd.Timestamp('2024-09-24'),pd.Timestamp('2024-10-24'),color='#f0efec',zorder=0)
# 阶段色带与标注（全部由 stages.S 自动生成）
cum=d.cum; est=d.index[d.est]
for _,rw in S.iterrows():
    ax.axvspan(pd.Timestamp(rw.start),pd.Timestamp(rw.end)+pd.Timedelta(days=1),color=RED if rw.GJD净额>0 else BLUE,alpha=0.10,lw=0,zorder=0)
top=max(cum.max(),0)*1.13+1500; bot=min(cum.min(),0)*1.13-1500
last={True:None,False:None}
for _,rw in S.iterrows():
    up=rw.GJD净额>0; s0,e=pd.Timestamp(rw.start),pd.Timestamp(rw.end); xm=s0+(e-s0)/2
    v=('≈' if rw.含估算 else '')+f'{rw.GJD净额:+,}'.replace('-','−') if abs(rw.GJD净额)>=1000 or rw.本次 else ''
    gap='\n' if last[up] is not None and (xm-last[up]).days<75 and not up else ''
    ax.text(xm,top-900 if up else bot+900,rw.编号+('本次' if rw.本次 else '')+'\n'+gap+v,ha='center',va='bottom' if up else 'top',fontsize=9.5,color=RED if up else BLUE)
    last[up]=xm
ax.set_ylim(bot,top)
if len(est):
    real=cum[:est[0]].iloc[:-1]
    ax.plot(real.index,real,color=INK,lw=2)
    ax.plot(cum.index[len(real)-1:],cum.iloc[len(real)-1:],color=INK,lw=2,ls=':')
    ax.plot(est,cum[est],'o',ms=8,mfc='#fcfcfb',mec=INK,mew=2,zorder=5)
    txt='、'.join(f'{t.month}/{t.day}' for t in est)+' 估算 '+'、'.join(f'{v:+.0f}' for v in d.gjd[est])+'\n累计 %.0f'%cum.iloc[-1]
    ax.annotate(txt,(est[-1],cum[est[-1]]),xytext=(-200,45),textcoords='offset points',fontsize=10,color=INK,arrowprops=dict(arrowstyle='-',color=INK2,lw=0.8))
else:
    ax.plot(cum.index,cum,color=INK,lw=2)
ax.text(pd.Timestamp('2024-09-26'),bot*0.3,'9/24–10/24\n散户主导\n（剔除）',fontsize=9,color=INK2)
ax.axhline(0,color=INK2,lw=0.8)
ax.set_ylabel('累计净流入（亿元）',color=INK2)
ax.set_title(f'国家队宽基ETF累计净流入（{d.index[0]:%Y/%-m/%-d}–{d.index[-1]:%Y/%-m/%-d}，23只宽基ETF，份额变化×均价）',loc='left',fontsize=14,color=INK,pad=34)
g=d.gjd
gr=g.where(~d.est,0)
bx.bar(d.index[gr>0],gr[gr>0],width=1.6,color=RED,label='国家队买入日')
bx.bar(d.index[gr<0],gr[gr<0],width=1.6,color=BLUE,label='国家队卖出日')
fol=d.follow
if (fol>0).any(): bx.bar(d.index[fol>0],fol[fol>0],bottom=gr[fol>0].clip(lower=0),width=1.6,color='#b4b2ab',label='跟风资金（已扣除）')
if len(est): bx.bar(est,g[est],width=1.6,color='#fcfcfb',edgecolor=RED,hatch='////',lw=0.8,label='估算（份额数据未出）')
bx.axhline(0,color=INK2,lw=0.8); bx.set_ylabel('单日（亿元）',color=INK2)
bx.legend(loc='lower left',frameon=False,fontsize=9.5,ncol=4)
bx.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1,4,7,10])); bx.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
fig.text(0.01,0.005,'识别规则：|当日份额×均价| > 0.3×常态成交额，且 |流入| / 超额成交 ≥ 0.4；常态成交额=前120日成交额20%分位；超额成交=成交额−常态。红=买入阶段 蓝=卖出阶段 灰=散户期。本次阶段已按成交结构扣除跟风资金。',fontsize=8.5,color=INK2)
fig.savefig(os.path.join(os.path.dirname(os.path.abspath(__file__)),'GJD_cumulative.png'),dpi=150,bbox_inches='tight',facecolor=fig.get_facecolor())
