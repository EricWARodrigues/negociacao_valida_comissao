import pandas as pd
from datetime import datetime

planilha = pd.read_excel('negociacao_valida_comissao\\Exemplo1.xlsx', sheet_name='planilha')
banco_de_dados = pd.read_excel('negociacao_valida_comissao\\Exemplo1.xlsx', sheet_name='banco_de_dados')
planilha.sort_values(['dt_negociacao'])
banco_de_dados = banco_de_dados.sort_values(['fatura'])

p = len(planilha['fatura'])
w = len(banco_de_dados['fatura'])

for pr in range(p):
    fat_planilha_atual = planilha['fatura'][pr]
    for wg in range(w):
        fat_banco_de_dados_atual = banco_de_dados['fatura'][wg]
        if fat_planilha_atual == fat_banco_de_dados_atual:
            planilha.loc[pr, 'dt_vencimento'] = banco_de_dados['dt_vencimento'][wg]
            
for pr in range(p):
    fat_planilha_atual = planilha['fatura'][pr]
    for wg in range(w):
        fat_banco_de_dados_atual = banco_de_dados['fatura'][wg]
        if fat_planilha_atual == fat_banco_de_dados_atual:
            planilha.loc[pr, 'dt_pagamento'] = banco_de_dados['dt_pagamento'][wg]

for pr in range(p):
    if (planilha['dt_negociacao'][pr] < planilha['dt_vencimento'][pr] or
        planilha['dt_negociacao'][pr] < planilha['dt_pagamento'][pr] or
        (planilha['dt_pagamento'][pr] - planilha['dt_negociacao'][pr]).days > 31):
        planilha.loc[pr, 'dt_neg_valida'] = True
    else:
        planilha.loc[pr, 'dt_neg_valida'] = False

dic_indices = {}
for i in range(p):
    cod_fatura = planilha['fatura'][i]
    if cod_fatura not in dic_indices:
        dic_indices[cod_fatura] = []
    dic_indices[cod_fatura].append(i)

from datetime import timedelta

planilha['venc_negociacao'] = None

for i in range(p):
    cod_fatura = planilha['fatura'][i]
    qtd_faturas_negociadas = len(dic_indices[cod_fatura])
    if qtd_faturas_negociadas < 2:
        planilha.at[i, 'venc_negociacao'] = planilha['dt_negociacao'][i] + timedelta(days=31)
    else:
        limite = dic_indices[cod_fatura][-1]
        if i == limite:
            planilha.at[i, 'venc_negociacao'] = planilha['dt_negociacao'][i] + timedelta(days=31)
        elif (planilha['dt_negociacao'][i + 1] - planilha['dt_negociacao'][i]).days < 8:
            planilha.at[i, 'venc_negociacao'] = planilha['dt_negociacao'][i] + timedelta(days=7)
        else:
            planilha.at[i, 'venc_negociacao'] = planilha['dt_negociacao'][i + 1] - timedelta(days=1)

for cod_fatura, indices in dic_indices.items():
    for i, idx in enumerate(indices):
        if i == 0:
            planilha.at[idx, 'dt_inicio_negociacao'] = planilha['dt_negociacao'][idx]
        else:
            planilha.at[idx, 'dt_inicio_negociacao'] = planilha.at[indices[i - 1], 'venc_negociacao'] + timedelta(days=1)
            
for i in range(p):
    if (planilha['dt_pagamento'][i] >= planilha['dt_inicio_negociacao'][i] and
        planilha['dt_pagamento'][i] <= planilha['venc_negociacao'][i]):
        planilha.loc[i, 'dt_neg_valida'] = True
    else:
        planilha.loc[i, 'dt_neg_valida'] = False

filename = 'Exemplo1_tratado'
writer = pd.ExcelWriter(f'.\\{filename}.xlsx')
planilha.to_excel(writer, index=False)
writer.close()