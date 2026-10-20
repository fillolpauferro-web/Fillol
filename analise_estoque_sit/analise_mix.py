"""Análise de mix (todos os distribuidores somados).

Uso: python analise_mix.py <arquivo.xlsx> [pasta_saida]
Decompõe a variação do valor do estoque em efeito volume, mix e preço e compara o mix do estoque com o do Sell Out por segmento, Curva ABC, família e SKU.
"""
import warnings; warnings.filterwarnings('ignore')
import pandas as pd,numpy as np,pickle
import sys,os
ARQ=sys.argv[1]; SAIDA=sys.argv[2] if len(sys.argv)>2 else 'saida_analise'; os.makedirs(SAIDA,exist_ok=True)
import analise_estoque_sit_base as base
d,L=base.carregar(ARQ)
M=['Abr','Mai','Jun','Jul','Ago']
cur=pd.concat([d[m][['SALES CODE','CURVA ABC']] for m in ['Ago','Jul','Jun','Mai']]).dropna().drop_duplicates('SALES CODE').set_index('SALES CODE')['CURVA ABC'].astype(str).replace({'0':'Sem curva','Others**':'Others'})
stat=L[L.Mes=='Ago'].groupby('SALES CODE')['STATUS LMID'].first()
seg=L.dropna(subset=['SEGMENTAÇÃO']).groupby('SALES CODE')['SEGMENTAÇÃO'].first().replace({'Bgx':'BGx'})
fam=L.groupby('SALES CODE')['FAMÍLIA'].agg(lambda s:s.mode()[0])
apr=L.groupby('SALES CODE')['APRESENTAÇÃO'].agg(lambda s:s.str.strip().mode()[0])
bu=L.groupby('SALES CODE')['BU LOCAL'].first()
L['curva']=L['SALES CODE'].map(cur).fillna('Sem curva'); L['status']=L['SALES CODE'].map(stat).fillna('Sem status')
L['segmento']=L['SALES CODE'].map(seg).fillna('Sem segmento'); L['familia']=L['SALES CODE'].map(fam); L['bu']=L['SALES CODE'].map(bu)
L['CDv']=L.STOCK_FLOOR_M1_CD*L.P; L['SOv']=L.SELLOUT_M1*L.P
F=['Mai','Jun','Jul','Ago']
def mix(dim):
    p=lambda v:L.pivot_table(index=dim,columns='Mes',values=v,aggfunc='sum',fill_value=0).reindex(columns=M,fill_value=0)
    u,v,so,sov=p('STOCK_FLOOR_M1_CD'),p('CDv'),p('SELLOUT_M1'),p('SOv')
    r=pd.DataFrame(index=u.index)
    r['un_abr'],r['un_ago'],r['rs_abr'],r['rs_ago']=u.Abr,u.Ago,v.Abr,v.Ago
    r['sh_un_abr']=u.Abr/u.Abr.sum(); r['sh_un_ago']=u.Ago/u.Ago.sum()
    r['sh_rs_abr']=v.Abr/v.Abr.sum(); r['sh_rs_ago']=v.Ago/v.Ago.sum()
    r['dsh_rs_pp']=(r.sh_rs_ago-r.sh_rs_abr)*100; r['dsh_un_pp']=(r.sh_un_ago-r.sh_un_abr)*100
    r['var_un_pct']=np.where(u.Abr>0,u.Ago/u.Abr-1,np.nan); r['var_rs']=v.Ago-v.Abr; r['var_rs_pct']=np.where(v.Abr>0,v.Ago/v.Abr-1,np.nan)
    r['so_rs_4m']=sov[F].sum(1); r['sh_so_rs']=r.so_rs_4m/r.so_rs_4m.sum()
    r['sh_so_rs_abr']=sov.Abr/sov.Abr.sum()
    r['gap_pp']=(r.sh_rs_ago-r.sh_so_rs)*100   # estoque acima (+) ou abaixo (-) do peso na demanda
    r['cob_ago']=u.Ago/so[F].mean(1).replace(0,np.nan)
    r['preco_medio_ago']=v.Ago/u.Ago.replace(0,np.nan)
    for m in M: r['sh_rs_'+m.lower()]=v[m]/v[m].sum()
    for m in M: r['sh_so_'+m.lower()]=sov[m]/sov[m].sum()
    return r.replace([np.inf,-np.inf],np.nan)
out={k:mix(k) for k in ['segmento','curva','bu','status','familia']}
sk=mix('SALES CODE'); sk.insert(0,'apresentacao',apr); sk.insert(1,'familia',fam); sk.insert(2,'curva',cur.reindex(sk.index).fillna('Sem curva')); sk.insert(3,'segmento',seg.reindex(sk.index).fillna('Sem segmento'))
out['sku']=sk
# value bridge
uA,uG=L[L.Mes=='Abr'].STOCK_FLOOR_M1_CD.sum(),L[L.Mes=='Ago'].STOCK_FLOOR_M1_CD.sum()
vA,vG=L[L.Mes=='Abr'].CDv.sum(),L[L.Mes=='Ago'].CDv.sum()
pA,pG=vA/uA,vG/uG
br=dict(rs_abr=vA,efeito_volume=(uG-uA)*pA,efeito_mix=uG*(pG-pA),efeito_preco=0.0,rs_ago=vG,preco_medio_abr=pA,preco_medio_ago=pG,un_abr=uA,un_ago=uG)
out['ponte']=br
# mix effect by family: contribution = (share_ago - share_abr)*uG*(p_f - pA)
fm=L.groupby(['familia','Mes']).agg(u=('STOCK_FLOOR_M1_CD','sum'),v=('CDv','sum')).unstack().fillna(0)
pf=(fm['v']['Ago']+fm['v']['Abr'])/(fm['u']['Ago']+fm['u']['Abr']).replace(0,np.nan)
contrib=((fm['u']['Ago']/uG-fm['u']['Abr']/uA)*uG*(pf-pA)).fillna(0)
out['familia']['contrib_mix']=contrib
# monthly avg price
out['mensal']=pd.DataFrame({'un':L.groupby('Mes').STOCK_FLOOR_M1_CD.sum(),'rs':L.groupby('Mes').CDv.sum(),'so_un':L.groupby('Mes').SELLOUT_M1.sum(),'so_rs':L.groupby('Mes').SOv.sum()}).reindex(M)
out['mensal']['preco_estoque']=out['mensal'].rs/out['mensal'].un; out['mensal']['preco_so']=out['mensal'].so_rs/out['mensal'].so_un

O=SAIDA+'/'
sk=out['sku'].copy(); br=out['ponte']
uA,uG,pA=br['un_abr'],br['un_ago'],br['preco_medio_abr']
sk['preco']=(sk.rs_ago+sk.rs_abr)/(sk.un_ago+sk.un_abr).replace(0,np.nan)
sk['preco']=sk.preco.fillna(0)
sk['ef_mix']=(sk.un_ago/uG-sk.un_abr/uA)*uG*(sk.preco-pA)
sk['ef_vol']=(sk.un_abr/uA)*(uG-uA)*sk.preco  # volume effect allocated pro-rata to Abr share
print('check mix',sk.ef_mix.sum(),br['efeito_mix'],'vol',sk.ef_vol.sum(),br['efeito_volume'])
# family/segment/curve contributions
for k,col in [('familia','familia'),('segmento','segmento'),('curva','curva')]:
    out[k]['ef_mix']=sk.groupby(col).ef_mix.sum()
# long dim table
rows=[]
for k,lab in [('segmento','Segmentação'),('curva','Curva ABC'),('status','Status')]:
    t=out[k].copy(); t=t[(t.un_abr+t.un_ago+t.so_rs_4m)>0]
    t.insert(0,'dimensao',lab); t.index.name='item'; rows.append(t.reset_index())
dim=pd.concat(rows)
order={'Curva ABC':['AAA','A','B','C','D','Others'],'Segmentação':['Expansion','Growth','Sustain','BGx'],'Status':['Active','Inactive','On hold']}
dim['ordem']=dim.apply(lambda r:order[r.dimensao].index(r['item']) if r['item'] in order[r.dimensao] else 9,axis=1)
dim=dim[dim.ordem<9]
cols=['dimensao','item','ordem','un_abr','un_ago','rs_abr','rs_ago','sh_rs_abr','sh_rs_ago','dsh_rs_pp','var_rs_pct','so_rs_4m','sh_so_rs','gap_pp','cob_ago','preco_medio_ago','ef_mix']
dim['ef_mix']=dim['ef_mix'].fillna(0) if 'ef_mix' in dim else 0
dim[cols].round(5).to_csv(O+'mix_dim.csv',index=False)
# monthly shares long (segmento,curva)
ml=[]
for k,lab in [('segmento','Segmentação'),('curva','Curva ABC')]:
    t=out[k]
    for it in order[lab]:
        if it not in t.index: continue
        for i,m in enumerate(['abr','mai','jun','jul','ago']):
            ml.append(dict(dimensao=lab,item=it,ordem=order[lab].index(it),mes=m.capitalize(),mes_ordem=i+1,sh_estoque=t.loc[it,'sh_rs_'+m],sh_sellout=t.loc[it,'sh_so_'+m]))
pd.DataFrame(ml).round(5).to_csv(O+'mix_mensal.csv',index=False)
f=out['familia'].reset_index().rename(columns={'index':'familia'})
f=f[(f.un_abr+f.un_ago+f.so_rs_4m)>0]
f[['familia','un_abr','un_ago','rs_abr','rs_ago','sh_rs_abr','sh_rs_ago','dsh_rs_pp','var_rs_pct','sh_so_rs','gap_pp','cob_ago','preco_medio_ago','ef_mix']].round(5).to_csv(O+'mix_familia.csv',index=False)
s=sk.reset_index().rename(columns={'SALES CODE':'sku'})
s=s[(s.un_abr+s.un_ago+s.so_rs_4m)>0]
s['rank']=s.rs_ago.rank(ascending=False,method='first')
s=s.sort_values('rs_ago',ascending=False); s['acum_rs']=s.rs_ago.cumsum()/s.rs_ago.sum()
s[['sku','apresentacao','familia','curva','segmento','un_abr','un_ago','rs_abr','rs_ago','sh_rs_ago','dsh_rs_pp','sh_so_rs','gap_pp','cob_ago','preco','ef_mix','acum_rs']].round(5).to_csv(O+'mix_sku.csv',index=False)
mm=out['mensal']
res=dict(rs_abr=br['rs_abr'],efeito_volume=br['efeito_volume'],efeito_mix=br['efeito_mix'],efeito_preco=0.0,rs_ago=br['rs_ago'],preco_estoque_abr=pA,preco_estoque_ago=br['preco_medio_ago'],preco_so_abr=mm.loc['Abr','preco_so'],preco_so_ago=mm.loc['Ago','preco_so'],
 skus_50=int((s.acum_rs<.5).sum()+1),skus_80=int((s.acum_rs<.8).sum()+1),skus_com_estoque=int((s.rs_ago>0).sum()))
pd.DataFrame([res]).round(4).to_csv(O+'mix_resumo.csv',index=False)
mm.reset_index().rename(columns={'Mes':'mes'}).assign(mes_ordem=range(1,6)).round(4).to_csv(O+'mix_preco_mensal.csv',index=False)

print(res); print(f.sort_values('ef_mix')[['familia','ef_mix','dsh_rs_pp','preco_medio_ago']].head(6)); print(f.sort_values('ef_mix')[['familia','ef_mix','dsh_rs_pp','preco_medio_ago']].tail(6))
print(dim[['dimensao','item','ef_mix','gap_pp','dsh_rs_pp']].round(2))
print(s.sort_values('gap_pp')[['sku','apresentacao','sh_rs_ago','sh_so_rs','gap_pp','cob_ago']].head(8).round(4).to_string())
print(s.sort_values('gap_pp')[['sku','apresentacao','sh_rs_ago','sh_so_rs','gap_pp','cob_ago']].tail(8).round(4).to_string())
