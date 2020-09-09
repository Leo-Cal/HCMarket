import xlwings as xw
import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt
from rm_separator import *

def rm_view_creator(book):
    wb = xw.Book(book)
    sht = wb.sheets['RM-View']
    sht.range('1:19965').value = ""
    sht = wb.sheets['RM']
    data = sht.range('A1').options(pd.DataFrame, expand='table').value
    sht = wb.sheets['RM-View']
    columns = data.columns
    setup = rm_separator(book,columns[0])
    #print(setup)
    indexes = setup.index
    #print(indexes)
    full_rm_df = pd.DataFrame([],index=indexes)

    full_rm_df[columns[i]] = rm_separator(book,columns[i])

    # Separate by size
    j=0
    small = full_rm_df[columns[j]].loc[(full_rm_df[columns[j]] >= 0.005) & (full_rm_df[columns[j]] < 0.01)]
    medium = full_rm_df[columns[j]].loc[(full_rm_df[columns[j]] >= 0.01) & (full_rm_df[columns[j]] < 0.05)]
    big = full_rm_df[columns[j]].loc[(full_rm_df[columns[j]] >= 0.05)]
    big = big.drop(index="")  # tira a linha de total

    #sht.range('A1').value = small
    #sht.range('D1').value = medium

    #print(result_join2)

