"""Leitura das 5 abas mensais e montagem da base longa (uma linha por cliente x SKU x mês)."""
def carregar(arq):
    import pickle, pandas as pd, numpy as np
    x=pd.read_excel(arq,sheet_name=None,header=1)
    M=['Abr','Mai','Jun','Jul','Ago']
    d={}
    for m,(k,df) in zip(M,x.items()):
        df=df.copy(); df.columns=[str(c).strip() for c in df.columns]
        df=df[df['CONCATENADO'].notna()]
        d[m]=df
    
    num=['STOCK_FLOOR_M1_CD','STOCK_FLOOR_M1_POS','TRANSITO','SELLOUT_M1','Sell in reconhecido mês atual','Cut off mês anterior','(+) Venda recon não entregue mês anterior','(+) Cut off antecipado mês atual','(-) Cut off antecipado mês anterior','Cut off mês atual','(-) Venda recon não entregue mês atual','SI indireto','Est. Chão Calculado','Variação Est. Chão Calculado','STOCK FLOOR_M1_ DC+POS','STOCK FLOOR_M2_ DC+POS','AVG SO 3M','Gross Price']
    dims=['CONCATENADO','GROUP_NAME','CUSTOMER_NAME','CUSTOMER_CHANNEL','FAMÍLIA','APRESENTAÇÃO','SALES CODE','SEGMENTAÇÃO','BU LOCAL','STATUS LMID','SIT?']
    L=[]
    for m,df in d.items():
        t=df[dims+[c for c in num if c in df]].copy()
        for c in num: t[c]=pd.to_numeric(t.get(c,0),errors='coerce').fillna(0)
        t['Mes']=m; L.append(t)
    long=pd.concat(L,ignore_index=True)
    long['FAMÍLIA']=long['FAMÍLIA'].str.upper().str.strip()
    long['SEGMENTAÇÃO']=long['SEGMENTAÇÃO'].str.title()
    # unify price per key (constant)
    price=long.groupby('CONCATENADO')['Gross Price'].max()
    long['P']=long['CONCATENADO'].map(price)
    long['SI_adj']=long['Sell in reconhecido mês atual']+long['Cut off mês anterior']+long['(+) Venda recon não entregue mês anterior']+long['(+) Cut off antecipado mês atual']
    long['DCPOS']=long['STOCK FLOOR_M1_ DC+POS']
    
    return d,long
