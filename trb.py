import os
import glob
import shutil
import pandas as pd
import numpy as np
from PyPDF2 import PdfReader
from data_extractor import extractText
if __name__ == '__main__':
    start_year = int(input('start year: '))
    end_year = int(input('end year: '))
    start_month = int(input('start month: '))
    end_month = int(input('end month: '))

    for year in range(start_year, end_year+1):
        for month in range(start_month, end_month+1):
            print(year, month)
            filename = str(year) + " " + str(month) + " TRB.xlsx"
            path = str(year) + " " + str(month) + "/*"
            for index, file in enumerate(glob.glob(path)):
                print(file)
                if index == 0:
                    sa = pd.read_excel(file, sheet_name='pivot_sa')
                    ca = pd.read_excel(file, sheet_name='pivot_ca')
                    fd = pd.read_excel(file, sheet_name='pivot_fd')
                    od = pd.read_excel(file, sheet_name='pivot_od')
                if index == 1:
                    securities = pd.read_excel(file, sheet_name='pivot_securities')
                    fund = pd.read_excel(file, sheet_name='pivot_fund')
                    bond = pd.read_excel(file, sheet_name='pivot_bond')
                if index == 2:
                    fitas = pd.read_excel(file, sheet_name='pivot_fitas')


            trb = pd.read_excel('trb.xlsx')
            trb = pd.merge(trb, sa, on=['CUSTCOD'], how='left')
            trb = pd.merge(trb, ca, on=['CUSTCOD'], how='left')
            trb = pd.merge(trb, fd, on=['CUSTCOD'], how='left')
            trb["Loans (4)"] = ""
            trb = pd.merge(trb, securities, on=['CUSTCOD'], how='left')
            trb = pd.merge(trb, fund, on=['CUSTCOD'], how='left')
            trb = pd.merge(trb, bond, on=['CUSTCOD'], how='left')
            trb["Total relationship (1)+(2)+(3)+(4)+(5)+(6)+(7)"] = ""
            trb = pd.merge(trb, od, on=['CUSTCOD'], how='left')


            #trb= trb[['CUSTCOD', 'Customer', 'CurBal/Owner(MOP)_x', 'CurBal/Owner(MOP)_y','CurBal/Owner(MOP)_x', 'Loans (4)' , 'outstanding per owner_x','outstanding per owner_y', 'outstanding per owner','Total relationship (1)+(2)+(3)+(4)+(5)+(6)+(7)', 'CurBal/Owner(MOP)_y']]
            trb.columns = ['CIF', 'Customer', "Saving Dep(1)", "Current Dep(2)", "Fixed Dep(3)","Loans (4)", "Securities(5)", "Funds(6)", "Bonds(7)", "Total relationship (1)+(2)+(3)+(4)+(5)+(6)+(7)", "OD Facility"]

            print(sa)
            print(ca)
            print(fd)
            print(od)
            print(securities)
            print(fund)
            print(bond)
            print(fitas)

            print(trb)
            with pd.ExcelWriter(filename) as writer:
                trb.to_excel(writer, 'result', index=False)

            sourcepath = filename
            despath = str(year) + " " + str(month) + "/" + filename
            shutil.move(filename, despath)