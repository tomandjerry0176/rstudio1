from model import *
EST1102=300.0
d.loc['2026-11-02','gjd']=EST1102
d['cum']=d.gjd.cumsum()
d['cum_user']=d.user50.cumsum()
pos={t:i for i,t in enumerate(d.index)}
gd=d.index[d.gjd_day|(d.index==pd.Timestamp('2026-11-02'))]
st=[];cur=[gd[0]]
for t in gd[1:]:
    if pos[t]-pos[cur[-1]]<=10 and np.sign(d.gjd[t])==np.sign(d.gjd[cur[-1]]): cur.append(t)
    else: st.append(cur);cur=[t]
st.append(cur)
rows=[]
for s in st:
    tot=d.gjd[s].sum()
    if abs(tot)>=500:
        seg=d.loc[s[0]:s[-1]]
        rows.append(dict(start=s[0].date(),end=s[-1].date(),交易日=len(seg),信号日=len(s),GJD净额=round(tot),
          日均=round(tot/len(s)),最大单日=round(d.gjd[s].abs().max()*np.sign(tot)),超额成交=round(seg.ex.sum()),
          流入_超额=round(seg.flow.sum()/seg.ex.sum(),2) if seg.flow.notna().all() else None, 均成交额=round(seg.amt.mean())))
S=pd.DataFrame(rows)

if __name__=='__main__':
    import os
    out=os.path.dirname(os.path.abspath(__file__))
    d[['amt','base','ex','flow','conc','gjd_day','gjd','cum','user50','cum_user']].round(3).to_csv(os.path.join(out,'gjd_daily.csv'),encoding='utf-8-sig')
    S.to_csv(os.path.join(out,'gjd_stages.csv'),index=False,encoding='utf-8-sig')
    print(S.to_string())
