from uf_separator import *
from rm_separator import *
from rm_view_creator import *
import xlwings as xw

def main():

    book = 'HCMarket.xlsx'
    wb = xw.Book(book)
    sht = wb.sheets('Summary')
    uf = sht.range('B5').value
    rm = "4211 RM Tubarão - núcleo metropolitano - SC"
    print("Starting...\n Book: %s \n UF: %s \n RM: %s" %(book,uf,rm))
    uf_separator('HCMarket.xlsx',uf)
    rm_separator('HCMarket.xlsx',rm)

    return 0

main()
