# Análise de estoque SIT por cliente (Abr–Ago 26)

Script: `analise_estoque_sit.py`. Lê a planilha com as 5 abas mensais (Sit Cust Abril … SIT Cust Ago) e gera CSVs.

```
python analise_estoque_sit.py Analise_SIT_Provisao_Out_26.xlsx saida
```

Lógica da reconciliação:

- **Simples**: Stock_Floor M1 CD de abril + Σ "Sell in reconhecido mês atual" − Σ SELLOUT_M1 (maio a agosto).
- **Ajustada** (igual à coluna "Est. Chão Calculado", encadeada): estoque CD+POS de abril + Sell In + "Cut off mês anterior" + "(+) Venda recon não entregue mês anterior" + "(+) Cut off antecipado mês atual" − Sell Out.

Observação: na planilha, "Variação Est. Chão Calculado" = estoque − |Est. calculado|, o que inverte o sinal quando o calculado é negativo. O script usa estoque − calculado.

A planilha de entrada não está versionada (dados de clientes).

## Análise de mix

```
python analise_mix.py Analise_SIT_Provisao_Out_26.xlsx saida
```

Soma todos os distribuidores. Como o Gross Price de cada SKU não muda entre os meses, a variação do valor do estoque se divide em:

- **Efeito volume** = (unidades ago − unidades abr) × preço médio de abril
- **Efeito mix** = unidades ago × (preço médio ago − preço médio abr); por SKU: Δ participação × unidades ago × (preço do SKU − preço médio de abril)
- **Efeito preço** = 0

"Estoque − demanda" = participação no estoque de agosto − participação no Sell Out de maio a agosto (em R$).
