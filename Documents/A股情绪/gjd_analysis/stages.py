from model import *

# ---- 份额数据还没出、只有成交额的最近几天：按近期国家队交易日各ETF的“流入/超额成交”比例估算 ----
last_flow=d.flow.last_valid_index()
miss=d.index[(d.index>last_flow)&d.amt.notna()]
d['est']=False
recent=d.loc[:last_flow].iloc[-30:]
rd=recent.index[recent.gjd_day]
if len(miss):
    if len(rd):
        ratio=(r.loc[rd].sum()/Aex.loc[rd].sum().replace(0,np.nan)).clip(-1,1).fillna(0)
        for t in miss:
            d.loc[t,'gjd']=float((ratio*Aex.loc[t]).sum()); d.loc[t,'est']=True
    else:
        print('提示：近30个交易日没有国家队交易日，缺份额数据的日期按0处理：',[t.date() for t in miss])
d['cum']=d.gjd.cumsum()
d['cum_user']=d.user50.cumsum()

# ---- 阶段划分：同方向、间隔<=10个交易日的国家队交易日连成一段；去掉开头孤立的零星小额信号（<100亿且与下一信号间隔>5天） ----
pos={t:i for i,t in enumerate(d.index)}
BUY='①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'; SELL='ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿⓀⓁⓂⓃⓄⓅⓆⓇⓈⓉ'
def build_stages():
    gd=d.index[d.gjd_day|d.est]
    st=[];cur=[gd[0]]
    for t in gd[1:]:
        if pos[t]-pos[cur[-1]]<=10 and np.sign(d.gjd[t])==np.sign(d.gjd[cur[-1]]): cur.append(t)
        else: st.append(cur);cur=[t]
    st.append(cur)
    rows=[]
    for s in st:
        t=s
        while len(t)>1 and pos[t[1]]-pos[t[0]]>5 and abs(d.gjd[t[0]])<100: t=t[1:]
        if abs(d.gjd[t].sum())>=0.9*abs(d.gjd[s].sum()): s=t   # 只在去掉的是零头时才裁剪
        tot=d.gjd[s].sum()
        if tot>=500 or tot<=-1000:   # 买入≥500亿、卖出≥1000亿才算一个阶段
            seg=d.loc[s[0]:s[-1]]; real=seg.flow.notna()
            rows.append(dict(start=s[0].date(),end=s[-1].date(),交易日=len(seg),信号日=len(s),GJD净额=round(tot),
              日均=round(tot/len(s)),最大单日=round(d.gjd[s].abs().max()*np.sign(tot)),超额成交=round(seg.ex.sum()),
              流入_超额=round(seg.flow[real].sum()/seg.ex[real].sum(),2), 均成交额=round(seg.amt.mean()),
              含估算=bool(seg.est.any()),跟风扣除=round(seg.follow.sum())))
    S=pd.DataFrame(rows)
    # 编号：买入 ①②③…，卖出 ⒶⒷⒸ…；最后一段若延续到最近5个交易日内则标“本次”
    nb=ns=0; lab=[]
    for _,rw in S.iterrows():
        if rw.GJD净额>0: lab.append(BUY[nb%20]); nb+=1
        else: lab.append(SELL[ns%20]); ns+=1
    S.insert(0,'编号',lab)
    S['本次']=False
    if len(S) and pos[pd.Timestamp(S.end.iloc[-1])]>=len(d)-6: S.loc[S.index[-1],'本次']=True
    return S

d['follow']=0.0
S=build_stages()

# ---- 本次阶段：从国家队交易日中扣除跟风资金 ----
# 超额成交 = 国家队(G) + 跟风(R)；流入 = C_PURE*G + PHI*R
#   C_PURE：纯国家队日的 流入/超额成交（⑤⑥阶段实测约0.9）
#   PHI：跟风资金的 净流入/超额成交，取本阶段及前20个交易日内“放量但单向性不足”的日子实测
# => 国家队流入 = C_PURE*(流入-PHI*超额成交)/(C_PURE-PHI)，其余记为跟风资金
C_PURE=0.9
PHI=None
if len(S) and S.本次.iloc[-1]:
    s0=pd.Timestamp(S.start.iloc[-1])
    w=d.iloc[max(pos[s0]-20,0):]
    fd=w.index[(~w.gjd_day)&(w.ex>0.3*w.base)&w.flow.notna()]
    PHI=float(np.clip(d.flow[fd].sum()/d.ex[fd].sum(),0,0.6)) if len(fd)>=3 else 0.2
    m=(d.index>=s0)&d.gjd_day&(d.flow>0)
    g=(C_PURE*(d.flow-PHI*d.ex)/(C_PURE-PHI)).clip(lower=0)
    g=np.minimum(g,d.flow)
    d.loc[m,'follow']=(d.flow-g)[m]; d.loc[m,'gjd']=g[m]
    d['cum']=d.gjd.cumsum()
    S=build_stages()

if __name__=='__main__':
    import os
    out=os.path.dirname(os.path.abspath(__file__))
    d[['amt','base','ex','flow','conc','gjd_day','est','follow','gjd','cum','user50','cum_user']].round(3).to_csv(os.path.join(out,'gjd_daily.csv'),encoding='utf-8-sig')
    S.to_csv(os.path.join(out,'gjd_stages.csv'),index=False,encoding='utf-8-sig')
    pd.set_option('display.width',250)
    print(S.to_string())
    print('\nPHI(跟风净流入/超额成交)=',PHI)
    print('\n最近10个交易日：'); print(d[['amt','ex','flow','conc','follow','gjd','est','cum']].iloc[-10:].round(1).to_string())
