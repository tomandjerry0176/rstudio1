import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import sys, logging
matplotlib.rcParams["axes.unicode_minus"]=False
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
fp="/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"; fm.fontManager.addfont(fp)
plt.rcParams["font.family"]=fm.FontProperties(fname=fp).get_name()
BLUE="#145B8C"; GOLD="#B8822A"; GRAY="#6B7280"; GRID="#E5E7EB"; LIGHT="#B9C3CC"
lab=[f"25/{m}" for m in range(1,13)]+[f"26/{m}" for m in range(1,10)]
val=[-4.1,-0.7,-3,0.5,0.4,0.3,-2.2,-1,-2.6,0.8,1.3,-3.6,-1.6,-1.2,0.5,-2.0,5.3,3.8,2.8,-1.5,-5.5]
col=[LIGHT]*12+[LIGHT]*4+[GOLD]*3+[BLUE]*2
fig,ax=plt.subplots(figsize=(10,3.6),dpi=300)
b=ax.bar(range(len(val)),val,color=col,width=0.68,zorder=3)
for i,v in enumerate(val):
    if i>=16: ax.text(i,v+(0.25 if v>0 else -0.25),f"{v:+.1f}",ha="center",va="bottom" if v>0 else "top",fontsize=10.5,color=GOLD if i<19 else BLUE)
ax.axhline(0,color="#9CA3AF",lw=0.9)
ax.set_xticks(range(len(val))); ax.set_xticklabels([l if (i%3==0 or i>=16) else "" for i,l in enumerate(lab)])
ax.set_ylim(-7,8.2); ax.set_yticks([-6,-3,0,3,6])
for s in ["top","right","left","bottom"]: ax.spines[s].set_visible(False)
ax.tick_params(colors=GRAY,labelsize=11,length=0); ax.grid(axis="y",color=GRID,lw=0.8); ax.set_axisbelow(True)
ax.text(0,1.04,"全金融机构其他贷款 月度净增减 万亿韩元",transform=ax.transAxes,fontsize=11,color=GRAY)
ax.text(0.005,0.9,"2025年单月最高 +1.3",transform=ax.transAxes,fontsize=10,color=GRAY)
ax.text(17,7.3,"5—7月合计 +11.9",ha="center",fontsize=10.5,color=GOLD)
ax.text(18.6,-5.0,"8—9月合计 -7.0",ha="right",fontsize=10.5,color=BLUE)
fig.tight_layout(); fig.savefig(sys.argv[1],dpi=300,facecolor="white")
