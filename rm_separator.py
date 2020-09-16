import xlwings as xw
import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

def rm_separator(book,rm):

    wb = xw.Book(book)
    sht = wb.sheets['RM-View']
    sht.range('1:200').value = ""
    sht = wb.sheets['RM']
    data = sht.range('A1').options(pd.DataFrame, expand='table').value
    indexes = data.index
    columns = data.columns

    path_small = 'exported data/'.join(rm)
    #print(path_small)
    #get data from RM
    rm_data = data[rm]
    #filter unimeds
    unimeds = rm_data.filter(like='UNIMED')  # pega lista de unimeds
    total_unimed = unimeds.sum()
    rm_data = rm_data[~rm_data.index.str.contains("UNIMED")]  # apaga UNIMEDS
    rm_data["000000-UNIMEDS"] = total_unimed  # coloca consolidado UNIMED
    rm_data = rm_data.sort_index()  # arruma a series

    mktshr = []
    for i in range(len(rm_data)):
        mktshr.append((rm_data[i] / rm_data[len(rm_data) - 1]))

    indexes_consolidado = rm_data.index.str[7:].values
    shr_rm_df = pd.DataFrame(data={rm:mktshr}, index=indexes_consolidado)

    small = shr_rm_df.loc[(shr_rm_df[rm] >= 0.005) & (shr_rm_df[rm] < 0.01)]
    medium = shr_rm_df.loc[(shr_rm_df[rm] >= 0.01) & (shr_rm_df[rm] < 0.05)]
    big = shr_rm_df.loc[(shr_rm_df[rm] >= 0.05)]
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
    total_caixas_small = caixas_small[rm].count() + autogestao_small[rm].count() + cooperativa_small[rm].count()
    total_caixas_medium = caixas_medium[rm].count() + autogestao_medium[rm].count() + cooperativa_medium[rm].count()
    total_caixas_big = caixas_big[rm].count() + autogestao_big[rm].count() + cooperativa_big[rm].count()
    operadoras_small = small[rm].count() - total_caixas_small
    operadoras_medium = medium[rm].count() - total_caixas_medium
    operadoras_big = big[rm].count() - total_caixas_big

    # Create the view
    sht = wb.sheets['RM-View']
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

    #Export data
    small.to_csv(path_or_buf='exported data/rm_small.txt',sep=';',index=True)
    medium.to_csv(path_or_buf='exported data/rm_medium.txt', sep=';', index=True)
    big.to_csv(path_or_buf='exported data/rm_big.txt', sep=';', index=True)

    return 0