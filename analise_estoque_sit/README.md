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
