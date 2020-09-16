import xlwings as xw
import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

def uf_separator(book,uf):

    wb = xw.Book(book)
    sht = wb.sheets['View']
    sht.range('1:200').value = ""
    sht = wb.sheets['UF']
    data = sht.range('A1').options(pd.DataFrame, expand='table').value
    indexes = data.index
    columns = data.columns

    #get data from RM
    uf_data = data[uf]
    #filter unimeds
    unimeds = uf_data.filter(like='UNIMED')  # pega lista de unimeds
    total_unimed = unimeds.sum()
    uf_data = uf_data[~uf_data.index.str.contains("UNIMED")]  # apaga UNIMEDS
    uf_data["000000-UNIMEDS"] = total_unimed  # coloca consolidado UNIMED
    uf_data = uf_data.sort_index()  # arruma a series

    mktshr = []
    for i in range(len(uf_data)):
        mktshr.append((uf_data[i] / uf_data[len(uf_data) - 1]))

    indexes_consolidado = uf_data.index.str[7:].values
    shr_uf_df = pd.DataFrame(data={uf:mktshr}, index=indexes_consolidado)

    small = shr_uf_df.loc[(shr_uf_df[uf] >= 0.005) & (shr_uf_df[uf] < 0.01)]
    medium = shr_uf_df.loc[(shr_uf_df[uf] >= 0.01) & (shr_uf_df[uf] < 0.05)]
    big = shr_uf_df.loc[(shr_uf_df[uf] >= 0.05)]
    big = big.drop(index="")  # tira a linha de total

    # Separa Caixas assistenciais/Autogestões
    caixas_small = small.filter(like='CAIXA', axis=0)
    autogestao_small = small.filter(like='AUTOGESTÃO', axis=0)
    cooperativa_small = small.filter(like='COOPERATIVA', axis=0)

    caixas_medium = medium.filter(like='CAIXA', axis=0)
    autogestao_medium = medium.filter(like='AUTOGESTÃO', axis=0)
    cooperativa_medium = medium.filter(like='COOPERATIVA', axis=0)

    caixas_big = big.filter(like='CAIXA', axis=0)
    autogestao_big = big.filter(like='AUTOGESTÃO', axis=0)
    cooperativa_big = big.filter(like='COOPERATIVA', axis=0)

    small_indexes = [k for k in small.index if ('CAIXA' in k or 'AUTOGESTÃO' in k or 'COOPERATIVA' in k or 'BENEFICENTE' in k)]
    df_operadoras_small = small.drop(small_indexes,axis=0)

    medium_indexes = [k for k in medium.index if ('CAIXA' in k or 'AUTOGESTÃO' in k or 'COOPERATIVA' in k or 'BENEFICENTE' in k)]
    df_operadoras_medium = medium.drop(medium_indexes, axis=0)

    big_indexes = [k for k in big.index if ('CAIXA' in k or 'AUTOGESTÃO' in k or 'COOPERATIVA' in k or 'BENEFICENTE' in k)]
    df_operadoras_big = big.drop(big_indexes, axis=0)

    total_caixas_small = caixas_small[uf].count() + autogestao_small[uf].count() + cooperativa_small[uf].count()
    total_caixas_medium = caixas_medium[uf].count() + autogestao_medium[uf].count() + cooperativa_medium[uf].count()
    total_caixas_big = caixas_big[uf].count() + autogestao_big[uf].count() + cooperativa_big[uf].count()
    operadoras_small = small[uf].count() - total_caixas_small
    operadoras_medium = medium[uf].count() - total_caixas_medium
    operadoras_big = big[uf].count() - total_caixas_big

    # Create the view
    sht = wb.sheets['View']
    sht.range('A1').value = "SMALL"
    sht.range('E1').value = "MEDIUM"
    sht.range('I1').value = "BIG"

    sht.range('A2').value = small
    sht.range('E2').value = medium
    sht.range('I2').value = big

    sht.range('N2').value = "Small"
    sht.range('N3').value = {"Caixas/Autogestão": total_caixas_small, "Operadoras ex-Unimed": operadoras_small}
    sht.range('P2').value = "Medium"
    sht.range('P3').value = {"Caixas/Autogestão": total_caixas_medium, "Operadoras ex-Unimed": operadoras_medium}
    sht.range('R2').value = "Big"
    sht.range('R3').value = {"Caixas/Autogestão": total_caixas_big, "Operadoras ex-Unimed": operadoras_big}

    sht.range('U1').value = "Small Operadoras"
    sht.range('X1').value = "Medium Operadoras"
    sht.range('AA1').value = "Big Operadoras"
    sht.range('U2').value = df_operadoras_small
    sht.range('X2').value = df_operadoras_medium
    sht.range('AA2').value = df_operadoras_big
    total_operadoras = pd.concat([df_operadoras_small,df_operadoras_medium,df_operadoras_big])
    total_operadoras["Members"] = uf_data[len(uf_data)-1] * total_operadoras[uf]
    print(total_operadoras)

    path = "exported data/%s_addressable.txt"%uf
    total_operadoras.to_csv(path_or_buf=path,sep=';',index=True)


    #Export data
    small.to_csv(path_or_buf='exported data/uf_small.txt',sep=';',index=True)
    medium.to_csv(path_or_buf='exported data/uf_medium.txt', sep=';', index=True)
    big.to_csv(path_or_buf='exported data/uf_big.txt', sep=';', index=True)

    return 0