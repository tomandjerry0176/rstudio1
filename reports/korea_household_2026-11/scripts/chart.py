"""Daily margin loans vs stock-pledged loans from the user's CSV (KOFIA FreeSIS).
Usage: python3 -I chart.py <csv> <out.png>"""
import sys, logging
import pandas as pd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as mdates
from matplotlib import font_manager as fm
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
fp="/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"; fm.fontManager.addfont(fp)
plt.rcParams["font.family"]=fm.FontProperties(fname=fp).get_name(); plt.rcParams["axes.unicode_minus"]=False
BLUE="#145B8C"; GOLD="#B8822A"; GRAY="#6B7280"; GRID="#E5E7EB"
df=pd.read_csv(sys.argv[1],encoding="utf-8-sig",dtype=str)
d=df[(df["频率"]=="日度")&(df["市场"]=="合计")&df["指标代码"].isin(["margin_balance","pledged_securities_balance"])].copy()
d["v"]=d["数值"].astype(float)/1e6; d["d"]=pd.to_datetime(d["统计期起"])
w=d.pivot_table(index="d",columns="指标代码",values="v")
fig,ax=plt.subplots(figsize=(10,3.6),dpi=300)
for d0,d1 in [("2026-07-14","2026-07-31"),("2026-10-13","2026-10-22")]:
    ax.axvspan(pd.Timestamp(d0),pd.Timestamp(d1),color=GOLD,alpha=0.10,lw=0)
ax.plot(w.index,w["margin_balance"],color=BLUE,lw=2.2)
ax.plot(w.index,w["pledged_securities_balance"],color=GOLD,lw=2.2)
m=w["margin_balance"]; p=w["pledged_securities_balance"]
def lab(s,dt,txt,col,off,ha="center"):
    ax.annotate(txt,(pd.Timestamp(dt),s[pd.Timestamp(dt)]),textcoords="offset points",xytext=off,ha=ha,va="center",fontsize=10,color=col)
lab(m,"2026-06-24","6/24 峰值 38.63",BLUE,(0,12)); lab(m,"2026-08-04","8/4 低点 27.40",BLUE,(8,-2),"left"); lab(m,"2026-10-12","10/12 34.29",BLUE,(0,12))
ax.text(w.index[-1],m.iloc[-1]-0.6,"信用融资 11/3 28.95",color=BLUE,fontsize=10.5,ha="right",va="top")
ax.text(w.index[-1],p.iloc[-1]-2.4,"存托证券质押融资 11/3 25.92",color=GOLD,fontsize=10.5,ha="right")
ax.set_ylim(20,41); ax.set_yticks([20,25,30,35,40])
ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m月"))
for s in ["top","right","left"]: ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color(GRID); ax.tick_params(colors=GRAY,labelsize=11,length=0)
ax.grid(axis="y",color=GRID,lw=0.8); ax.set_axisbelow(True)
ax.text(0,1.04,"证券公司信用余额 万亿韩元 日度",transform=ax.transAxes,fontsize=11,color=GRAY)
ax.text(0,-0.17,"金融投资协会FreeSIS；阴影为7月、10月急跌期",transform=ax.transAxes,fontsize=9.5,color=GRAY)
fig.tight_layout(); fig.savefig(sys.argv[2],dpi=300,facecolor="white")
