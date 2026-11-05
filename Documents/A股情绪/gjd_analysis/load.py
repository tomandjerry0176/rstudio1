import pandas as pd, numpy as np
import os
F=os.path.join(os.path.dirname(os.path.abspath(__file__)),'..','GJD ETF份额.xlsx')
def load():
    a=pd.read_excel(F,'成交额',header=None,skiprows=3)
    a=a[pd.to_datetime(a[0],errors='coerce').notna()]
    amt=pd.Series(pd.to_numeric(a[24]).values,index=pd.to_datetime(a[0]),name='amt')
    b=pd.read_excel(F,'份额 均价',header=None,skiprows=3)
    b=b[pd.to_datetime(b[0],errors='coerce').notna()]
    flow=pd.Series(pd.to_numeric(b[24],errors='coerce').values,index=pd.to_datetime(b[0]),name='flow')
    nb=pd.Series(pd.to_numeric(b[26],errors='coerce').values,index=pd.to_datetime(b[0]),name='nb')
    out=pd.concat([amt,flow,nb],axis=1)
    return out[out.amt.notna()]   # 去掉还没有成交额的空日期行

def raw():
    x=pd.read_excel(F,'份额 均价fml',header=None)
    hdr=x.iloc[3].tolist()
    x=x.iloc[4:]; x=x[pd.to_datetime(x[0],errors='coerce').notna()]
    idx=pd.to_datetime(x[0])
    sh={hdr[i]:pd.to_numeric(x[i],errors='coerce').values for i in range(1,29)}
    pr={hdr[i]:pd.to_numeric(x[i],errors='coerce').values for i in range(31,60)}
    codes=[c for c in hdr[1:24]]
    flows=pd.DataFrame({c:sh[c]*pr[c] for c in codes},index=idx)
    return flows
def nf():
    x=pd.read_excel(F,'净流入fml',header=None)
    hdr=x.iloc[3].tolist(); x=x.iloc[4:]; x=x[pd.to_datetime(x[0],errors='coerce').notna()]
    df=pd.DataFrame({hdr[i]:pd.to_numeric(x[i],errors='coerce').values for i in range(1,24)},index=pd.to_datetime(x[0]))
    return df
