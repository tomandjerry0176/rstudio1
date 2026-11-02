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
# stage shading
lab={}
names={0:'①',1:'②',2:'③',3:'④',4:'⑤',5:'',6:'',7:'⑥',8:'Ⓐ',9:'',10:'Ⓑ',11:'Ⓒ',12:'⑦',13:'',14:'⑧ 本次'}
for i,rw in S.iterrows():
    c=RED if rw.GJD净额>0 else BLUE
    s0,s1=pd.Timestamp(rw.start),pd.Timestamp(rw.end)
    if i==14: s0=pd.Timestamp('2026-10-09')
    ax.axvspan(s0,s1+pd.Timedelta(days=1),color=c,alpha=0.10,lw=0,zorder=0)
cum=d.cum
ax.plot(d.index,d.cum_user,color=GRAY,lw=1.4,ls='--',label='原方法：|流入|>50 硬阈值（不含11/2）')
ax.plot(cum.index[:-1],cum[:-1],color=INK,lw=2,label='国家队累计净流入（本文识别）')
ax.plot(cum.index[-2:],cum[-2:],color=INK,lw=2,ls=':')
ax.plot(cum.index[-1],cum.iloc[-1],'o',ms=8,mfc='#fcfcfb',mec=INK,mew=2,zorder=5)
ax.annotate('11/2 估算 +300\n累计 %.0f'%cum.iloc[-1],(cum.index[-1],cum.iloc[-1]),xytext=(-200,45),textcoords='offset points',fontsize=10,color=INK,arrowprops=dict(arrowstyle='-',color=INK2,lw=0.8))
ann={0:('①','',1),1:('②','',1),2:('③','+3,303',1),3:('④','',1),4:('⑤','+3,578',1),7:('⑥','+2,225',1),
     8:('Ⓐ','−9,652',0),10:('Ⓑ','−3,505',0),11:('Ⓒ','\n−2,010',0),12:('⑦','+1,406',1),14:('⑧本次','≈+2,580',1)}
for i,(k,v,up) in ann.items():
    rw=S.loc[i]; s0=pd.Timestamp('2026-10-09') if i==14 else pd.Timestamp(rw.start); e=pd.Timestamp(rw.end)
    xm=s0+(e-s0)/2
    ax.text(xm,11300 if up else -8300,k+'\n'+v,ha='center',va='bottom' if up else 'top',fontsize=9.5,color=RED if up else BLUE)
ax.set_ylim(-9600,12600)
ax.text(pd.Timestamp('2024-09-26'),-2500,'9/24–10/24\n散户主导\n（剔除）',fontsize=9,color=INK2)
ax.axhline(0,color=INK2,lw=0.8)
ax.set_ylabel('累计净流入（亿元）',color=INK2)
ax.set_title('国家队宽基ETF累计净流入（2023/8/1 起，23只宽基ETF，份额变化×均价）',loc='left',fontsize=14,color=INK,pad=34)
ax.legend(loc='lower left',frameon=False,fontsize=10)
g=d.gjd
bx.bar(d.index[g>0],g[g>0],width=1.6,color=RED,label='国家队买入日')
bx.bar(d.index[g<0],g[g<0],width=1.6,color=BLUE,label='国家队卖出日')
bx.bar(d.index[-1:],g[-1:],width=1.6,color='#fcfcfb',edgecolor=RED,hatch='////',lw=0.8,label='11/2 估算')
bx.axhline(0,color=INK2,lw=0.8); bx.set_ylabel('单日（亿元）',color=INK2)
bx.legend(loc='lower left',frameon=False,fontsize=9.5,ncol=3)
bx.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1,4,7,10])); bx.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
fig.text(0.01,0.005,'识别规则：|当日份额×均价| > 0.3×常态成交额，且 |流入| / 超额成交 ≥ 0.4；常态成交额=前120日成交额20%分位；超额成交=成交额−常态。红=买入阶段 蓝=卖出阶段 灰=散户期。11/2 份额数据缺失，按本阶段各ETF“流入/超额成交”比例估算。',fontsize=8.5,color=INK2)
fig.savefig(os.path.join(os.path.dirname(os.path.abspath(__file__)),'GJD_cumulative.png'),dpi=150,bbox_inches='tight',facecolor=fig.get_facecolor())
