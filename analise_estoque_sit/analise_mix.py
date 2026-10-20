"""Análise de mix em volume (unidades), todos os distribuidores somados.

Uso: python analise_mix.py <arquivo.xlsx> [pasta_saida]
Estoque CD e Sell Out em unidades por Curva ABC, Segmentação, família e SKU; cobertura e peso no estoque vs peso na demanda.
"""
import warnings; warnings.filterwarnings('ignore')
import pandas as pd,numpy as np
import sys,os
ARQ=sys.argv[1]; SAIDA=sys.argv[2] if len(sys.argv)>2 else 'saida_analise'; os.makedirs(SAIDA,exist_ok=True); O=SAIDA+'/'
import analise_estoque_sit_base as base
d,L=base.carregar(ARQ)
M=['Abr','Mai','Jun','Jul','Ago']; F=['Mai','Jun','Jul','Ago']
cur=pd.concat([d[m][['SALES CODE','CURVA ABC']] for m in ['Ago','Jul','Jun','Mai']]).dropna().drop_duplicates('SALES CODE').set_index('SALES CODE')['CURVA ABC'].astype(str).replace({'0':'Sem curva','Others**':'Others'})
seg=L.dropna(subset=['SEGMENTAÇÃO']).groupby('SALES CODE')['SEGMENTAÇÃO'].first().replace({'Bgx':'BGx'})
fam=L.groupby('SALES CODE')['FAMÍLIA'].agg(lambda s:s.mode()[0]); apr=L.groupby('SALES CODE')['APRESENTAÇÃO'].agg(lambda s:s.str.strip().mode()[0])
L['curva']=L['SALES CODE'].map(cur).fillna('Sem curva'); L['segmento']=L['SALES CODE'].map(seg).fillna('Sem segmento'); L['familia']=L['SALES CODE'].map(fam)
order={'Curva ABC':['AAA','A','B','C','D','Others'],'Segmentação':['Expansion','Growth','Sustain','BGx']}
def tab(col):
    p=lambda v:L.pivot_table(index=col,columns='Mes',values=v,aggfunc='sum',fill_value=0).reindex(columns=M,fill_value=0)
    u,so=p('STOCK_FLOOR_M1_CD'),p('SELLOUT_M1')
    r=pd.DataFrame(index=u.index)
    for m in M: r['un_'+m.lower()]=u[m]
    for m in M: r['so_'+m.lower()]=so[m]
    r['var_un']=u.Ago-u.Abr; r['var_pct']=np.where(u.Abr>0,u.Ago/u.Abr-1,np.nan)
    r['so_medio']=so[F].mean(1); r['var_so_pct']=np.where(so.Abr>0,so.Ago/so.Abr-1,np.nan)
    r['cob_abr']=u.Abr/so.Abr.replace(0,np.nan); r['cob_ago']=u.Ago/r.so_medio.replace(0,np.nan)
    r['sh_un_abr']=u.Abr/u.Abr.sum(); r['sh_un_ago']=u.Ago/u.Ago.sum(); r['sh_so']=so[F].sum(1)/so[F].sum().sum()
    r['gap_pp']=(r.sh_un_ago-r.sh_so)*100
    return r.replace([np.inf,-np.inf],np.nan)
dims=[];mens=[]
for lab,col in [('Curva ABC','curva'),('Segmentação','segmento')]:
    t=tab(col); t=t[t.index.isin(order[lab])]
    t['ordem']=[order[lab].index(i) for i in t.index]; t.insert(0,'dimensao',lab); t.index.name='item'
    dims.append(t.reset_index())
    for it in t.index:
        for i,m in enumerate(M): mens.append(dict(dimensao=lab,item=it,ordem=order[lab].index(it),mes=m,mes_ordem=i+1,estoque_un=t.loc[it,'un_'+m.lower()],sellout_un=t.loc[it,'so_'+m.lower()]))
dim=pd.concat(dims).sort_values(['dimensao','ordem']); dim.round(4).to_csv(O+'mixu_dim.csv',index=False)
pd.DataFrame(mens).round(1).to_csv(O+'mixu_mensal.csv',index=False)
f=tab('familia'); f.index.name='familia'; f=f[(f[['un_abr','un_ago','so_medio']].sum(1))>0].reset_index(); f.round(4).to_csv(O+'mixu_familia.csv',index=False)
s=tab('SALES CODE'); s.index.name='sku'; s.insert(0,'apresentacao',apr); s.insert(1,'familia',fam); s.insert(2,'curva',cur.reindex(s.index).fillna('Sem curva')); s.insert(3,'segmento',seg.reindex(s.index).fillna('Sem segmento'))
s=s[(s[['un_abr','un_ago','so_medio']].sum(1))>0].sort_values('un_ago',ascending=False); s['acum_un']=s.un_ago.cumsum()/s.un_ago.sum()
s.reset_index().round(4).to_csv(O+'mixu_sku.csv',index=False)
tot=dict(un_abr=L[L.Mes=='Abr'].STOCK_FLOOR_M1_CD.sum(),un_ago=L[L.Mes=='Ago'].STOCK_FLOOR_M1_CD.sum(),so_abr=L[L.Mes=='Abr'].SELLOUT_M1.sum(),so_medio=L[L.Mes.isin(F)].groupby('Mes').SELLOUT_M1.sum().mean(),
 skus_50=int((s.acum_un<.5).sum()+1),skus_80=int((s.acum_un<.8).sum()+1),skus_com_estoque=int((s.un_ago>0).sum()))
tot['var_un']=tot['un_ago']-tot['un_abr']; tot['var_pct']=tot['un_ago']/tot['un_abr']-1; tot['cob_ago']=tot['un_ago']/tot['so_medio']; tot['cob_abr']=tot['un_abr']/tot['so_abr']
pd.DataFrame([tot]).round(4).to_csv(O+'mixu_resumo.csv',index=False)
