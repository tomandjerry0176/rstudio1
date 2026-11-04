import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import datetime as dt, sys
plt.rcParams["axes.unicode_minus"]=False
fp="/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
fm.fontManager.addfont(fp)
plt.rcParams["font.family"]=fm.FontProperties(fname=fp).get_name()
BLUE="#145B8C"; GOLD="#B8822A"; GRAY="#6B7280"; GRID="#E5E7EB"
pts=[("2026-06-24",38.6328,"6/24 峰值 38.63"),("2026-07-16",33.3624,None),("2026-07-31",28.9,None),
     ("2026-08-04",27.4038,"8/4 低点 27.40"),("2026-08-26",33.1023,None),("2026-10-12",34.2887,"10/12 34.29"),
     ("2026-10-16",32.2528,None),("2026-10-22",30.3633,None),("2026-10-30",29.0664,"10/30 29.07"),("2026-11-02",29.216,None)]
x=[dt.date.fromisoformat(a) for a,_,_ in pts]; y=[b for _,b,_ in pts]
fig,ax=plt.subplots(figsize=(10,3.6),dpi=300)
ax.plot(x,y,color=BLUE,lw=2.4,marker="o",ms=5,zorder=3)
for (a,b,l),xi in zip(pts,x):
    if l:
        off,ha={"低点":((8,-4),"left"),"10/30":((0,-18),"center")}.get(next((k for k in ["低点","10/30"] if k in l),None),((0,9),"center"))
        ax.annotate(l,(xi,b),textcoords="offset points",xytext=off,ha=ha,va="center" if "低点" in l else "bottom",fontsize=10.5,color=BLUE)
for d0,d1,lab in [("2026-07-14","2026-07-30","7月急跌\n两天连续熔断"),("2026-10-13","2026-10-22","10月再跌\n10/15熔断")]:
    ax.axvspan(dt.date.fromisoformat(d0),dt.date.fromisoformat(d1),color=GOLD,alpha=0.13,lw=0)
    ax.text(dt.date.fromisoformat(d0)+(dt.date.fromisoformat(d1)-dt.date.fromisoformat(d0))/2,24.6,lab,ha="center",va="bottom",fontsize=9.5,color=GOLD)
ax.set_ylim(24,41); ax.set_yticks([25,30,35,40])
ax.set_xlim(dt.date(2026,6,15),dt.date(2026,11,8))
import matplotlib.dates as mdates
ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m月"))
for s in ["top","right","left"]: ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color(GRID); ax.tick_params(colors=GRAY,labelsize=11,length=0)
ax.grid(axis="y",color=GRID,lw=0.8); ax.set_axisbelow(True)
ax.text(0,1.04,"信用融资余额 万亿韩元",transform=ax.transAxes,fontsize=11,color=GRAY)
ax.text(0,-0.17,"金融投资协会口径，媒体报道的个别交易日读数，点间未连续观测",transform=ax.transAxes,fontsize=9.5,color=GRAY)
fig.tight_layout(); fig.savefig(sys.argv[1],dpi=300,facecolor="white")
