"""Análise de estoque SIT por cliente (abas Abril a Agosto).

Uso:  python analise_estoque_sit.py <arquivo.xlsx> [pasta_saida]

1. Variação Abr -> Ago do Stock_Floor M1 CD em unidades e R$ (Gross Price), por grupo, família e SKU.
2. Reconciliação: estoque de abril + Sell In - Sell Out (simples) e com os ajustes de cut-off
   (mesma lógica da coluna Est. Chão Calculado), comparado ao estoque real de agosto.
3. Alertas de ruptura/excesso por cobertura e checagens de qualidade.
Gera CSVs na pasta de saída.
"""
import sys, os, warnings; warnings.filterwarnings('ignore')
ARQ=sys.argv[1]; SAIDA=sys.argv[2] if len(sys.argv)>2 else 'saida_analise'
os.makedirs(SAIDA,exist_ok=True)
import pickle, pandas as pd, numpy as np
x=pd.read_excel(ARQ,sheet_name=None,header=1)
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

L=long

import pandas as pd,numpy as np
M=['Abr','Mai','Jun','Jul','Ago']; F=M[1:]
L['CDv']=L['STOCK_FLOOR_M1_CD']*L['P']
def agg(keys):
    p=lambda v: L.pivot_table(index=keys,columns='Mes',values=v,aggfunc='sum',fill_value=0).reindex(columns=M,fill_value=0)
    cd,cdv,dp,so,si,sia,tr=p('STOCK_FLOOR_M1_CD'),p('CDv'),p('DCPOS'),p('SELLOUT_M1'),p('Sell in reconhecido mês atual'),p('SI_adj'),p('TRANSITO')
    sov=L.assign(v=L.SELLOUT_M1*L.P).pivot_table(index=keys,columns='Mes',values='v',aggfunc='sum',fill_value=0).reindex(columns=M,fill_value=0)
    r=pd.DataFrame(index=cd.index)
    for m in M: r[f'CD {m}']=cd[m]
    r['Δ un']=cd.Ago-cd.Abr; r['Δ %']=np.where(cd.Abr>0,(cd.Ago/cd.Abr-1),np.nan)
    for m in M: r[f'CD R$ {m}']=cdv[m]
    r['Δ R$']=cdv.Ago-cdv.Abr; r['Δ R$ %']=np.where(cdv.Abr>0,cdv.Ago/cdv.Abr-1,np.nan)
    r['SI Mai-Ago']=si[F].sum(1); r['SO Mai-Ago']=so[F].sum(1); r['Ajustes cut-off Mai-Ago']=(sia-si)[F].sum(1)
    r['Calc simples Ago']=cd.Abr+r['SI Mai-Ago']-r['SO Mai-Ago']
    r['Dif simples (Real-Calc)']=cd.Ago-r['Calc simples Ago']
    r['Calc ajustado Ago (CD+POS)']=dp.Abr+sia[F].sum(1)-r['SO Mai-Ago']
    r['Real Ago (CD+POS)']=dp.Ago
    r['Dif ajustada (Real-Calc)']=dp.Ago-r['Calc ajustado Ago (CD+POS)']
    r['Dif ajustada %']=np.where(r['Real Ago (CD+POS)']!=0,r['Dif ajustada (Real-Calc)']/r['Real Ago (CD+POS)'],np.nan)
    pr=(cdv.Ago/cd.Ago.replace(0,np.nan)).fillna((cdv.Abr/cd.Abr.replace(0,np.nan)))
    r['Dif ajustada R$']=r['Dif ajustada (Real-Calc)']*pr.fillna(0)
    r['Cobertura Ago (meses)']=cd.Ago/(so[F].mean(1).replace(0,np.nan))
    r['Cobertura Abr (meses)']=cd.Abr/(so.Abr.replace(0,np.nan))
    r['SO médio Mai-Ago']=so[F].mean(1); r['SO R$ Mai-Ago']=sov[F].sum(1)
    r['Trânsito Ago']=tr.Ago
    return r,(cd,dp,so,si,sia,tr)

O=SAIDA+'/'
ren={'CD Abr':'cd_abr','CD Mai':'cd_mai','CD Jun':'cd_jun','CD Jul':'cd_jul','CD Ago':'cd_ago','Δ un':'var_un','Δ %':'var_pct','CD R$ Abr':'rs_abr','CD R$ Ago':'rs_ago','Δ R$':'var_rs','Δ R$ %':'var_rs_pct','SI Mai-Ago':'si','SO Mai-Ago':'so','Ajustes cut-off Mai-Ago':'cutoff','Calc simples Ago':'calc_simples','Dif simples (Real-Calc)':'dif_simples','Calc ajustado Ago (CD+POS)':'calc_ajust','Real Ago (CD+POS)':'real_ago','Dif ajustada (Real-Calc)':'dif_ajust','Dif ajustada %':'dif_ajust_pct','Dif ajustada R$':'dif_ajust_rs','Cobertura Ago (meses)':'cob_ago','Cobertura Abr (meses)':'cob_abr','SO médio Mai-Ago':'so_medio','SO R$ Mai-Ago':'so_rs','Trânsito Ago':'transito_ago'}
keep=['cd_abr','cd_mai','cd_jun','cd_jul','cd_ago','var_un','var_pct','rs_abr','rs_ago','var_rs','var_rs_pct','si','so','cutoff','calc_simples','dif_simples','calc_ajust','real_ago','dif_ajust','dif_ajust_pct','dif_ajust_rs','cob_abr','cob_ago','transito_ago']
def R(df,dec=4):
    return df.replace([np.inf,-np.inf],np.nan).round(dec)
rg,_=agg(['GROUP_NAME']); g=rg.rename(columns=ren)[keep].reset_index().rename(columns={'GROUP_NAME':'grupo'})
g=g[g.grupo!='DCENTER']  # 227 un em Abr, zerado depois
R(g).to_csv(O+'grupos.csv',index=False)
# monthly
L['res']=np.where(L.Mes=='Abr',np.nan,L['DCPOS']-(L['STOCK FLOOR_M2_ DC+POS']+L['SI_adj']-L['SELLOUT_M1']))
L['E']=L['Est. Chão Calculado']
m=L.groupby('Mes').agg(cd=('STOCK_FLOOR_M1_CD','sum'),cd_rs=('CDv','sum'),cd_pos=('DCPOS','sum'),transito=('TRANSITO','sum'),so=('SELLOUT_M1','sum'),si=('Sell in reconhecido mês atual','sum'),si_ajust=('SI_adj','sum'),residual=('res','sum'),var_planilha=('Variação Est. Chão Calculado','sum'),linhas_neg=('E',lambda s:(s<0).sum()))
m['var_correta']=L.assign(v=L.DCPOS-L.E).groupby('Mes')['v'].sum()
m['so_rs']=L.assign(v=L.SELLOUT_M1*L.P).groupby('Mes')['v'].sum()
m=m.reindex(M); m.loc['Abr','residual']=np.nan
m['mes_ordem']=range(1,6); m=m.reset_index().rename(columns={'Mes':'mes'})
R(m,1).to_csv(O+'mensal.csv',index=False)
rr=L[L.Mes!='Abr'].groupby(['GROUP_NAME','Mes'])['res'].sum().reset_index().rename(columns={'GROUP_NAME':'grupo','Mes':'mes','res':'residual'})
rr=rr[rr.grupo!='DCENTER']; rr['mes_ordem']=rr.mes.map({k:i for i,k in enumerate(M)})
R(rr,1).sort_values(['grupo','mes_ordem']).to_csv(O+'residuo_grupo.csv',index=False)
rf,_=agg(['FAMÍLIA']); f=rf.rename(columns=ren)[keep].reset_index().rename(columns={'FAMÍLIA':'familia'})
R(f).to_csv(O+'familias.csv',index=False)
rgf,_=agg(['GROUP_NAME','FAMÍLIA']); gf=rgf.rename(columns=ren)[['cd_abr','cd_ago','var_un','var_pct','rs_abr','rs_ago','var_rs','si','so','cutoff','dif_ajust','dif_ajust_rs','cob_ago']].reset_index().rename(columns={'GROUP_NAME':'grupo','FAMÍLIA':'familia'})
gf=gf[(gf[['cd_abr','cd_ago','si','so']].abs().sum(1)>0)&(gf.grupo!='DCENTER')]
R(gf,3).to_csv(O+'grupo_familia.csv',index=False)
rs,_=agg(['GROUP_NAME','SALES CODE','APRESENTAÇÃO']); s=rs.rename(columns=ren).reset_index().rename(columns={'GROUP_NAME':'grupo','SALES CODE':'sku','APRESENTAÇÃO':'apresentacao','SO médio Mai-Ago':'so_medio'})
s=s.rename(columns={'SO médio Mai-Ago':'so_medio','SO R$ Mai-Ago':'so_rs'})
rup=s[(s.so_medio>1000)&(s.cob_ago<0.5)].assign(tipo='Risco de ruptura')
exc=s[(s.cob_ago>4)&(s.rs_ago>100000)].assign(tipo='Excesso')
sk=pd.concat([rup,exc])[['tipo','grupo','sku','apresentacao','cd_ago','so_medio','cob_ago','rs_ago','so_rs']]
sk['apresentacao']=sk.apresentacao.str.strip()
R(sk,2).sort_values(['tipo','so_rs'],ascending=[False,False]).to_csv(O+'skus_alerta.csv',index=False)
# status
A=L[L.Mes=='Ago']; st=A.groupby('STATUS LMID').agg(un=('STOCK_FLOOR_M1_CD','sum'),rs=('CDv','sum'),linhas=('CONCATENADO','count')).reset_index().rename(columns={'STATUS LMID':'status'})
R(st,1).to_csv(O+'status_ago.csv',index=False)
print(g[['grupo','dif_ajust']].to_string()); print(m.to_string()); print(len(gf),len(sk),rup.shape,exc.shape)

