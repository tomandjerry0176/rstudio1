from load import *
import warnings; warnings.filterwarnings('ignore')
d=load(); r=raw(); d['flow']=r.sum(axis=1,min_count=1).reindex(d.index)
d['base']=d.amt.rolling(120,min_periods=40).quantile(0.2).shift(1).bfill()
d['ex']=(d.amt-d.base).clip(lower=0)
f=d.flow
med=f.rolling(120,min_periods=40).median().shift(1).bfill()
mad=(f-med).abs().rolling(120,min_periods=40).median().shift(1).bfill()*1.4826
d['T']=np.maximum(2.5*mad,30)
d['conc']=f.abs()/d.ex.replace(0,np.nan)
W=(d.index>='2024-09-24')&(d.index<='2024-10-24')
d['retail']=W
d['gjd_day']=(f.abs()>0.3*d.base)&(d.conc>=0.4)&(~W)
d['gjd']=np.where(d.gjd_day,f,0.0)
d['user50']=np.where((f.abs()>50)&(~W),f,0.0)

a=pd.read_excel(F,'成交额',header=None)
codes=a.iloc[2,1:24].tolist()
a=a.iloc[3:]; a=a[pd.to_datetime(a[0],errors='coerce').notna()]
A=pd.DataFrame({c:pd.to_numeric(a[i+1],errors='coerce').values for i,c in enumerate(codes)},index=pd.to_datetime(a[0])).reindex(d.index)
Abase=A.rolling(120,min_periods=40).quantile(0.2).shift(1).bfill()
Aex=(A-Abase).clip(lower=0)
