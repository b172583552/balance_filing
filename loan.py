import os
import glob
import shutil
import pandas as pd
import numpy as np
from PyPDF2 import PdfReader
from data_extractor import extractText
#for i in range(1,66):
    #os.mkdir('2021 0' + str(i))

month = 12
for year in range(2017,2021):
    mortgage = pd.DataFrame()
    termloan = pd.DataFrame()
    odlimit = pd.read_excel('2) Loan by CUSTOMER.xlsx', sheet_name=2)
    fitas = pd.read_excel('2) Loan by CUSTOMER.xlsx', sheet_name=4)
    bookrate = pd.read_excel('2) Loan by CUSTOMER.xlsx', sheet_name=5)
    filename = str(year)+" "+str(month)+' loan.xlsx'

    required_odlimit = odlimit[(odlimit.Year == year) & (odlimit.Month == month)]
    required_fitas = fitas[(fitas.Year == year) & (fitas.Month == month)]

    pivot_odlimit = pd.pivot_table(required_odlimit, index='CUSTCOD', values='Principle/Owner (Limit in MOP)', aggfunc=np.sum)
    pivot_fitas = pd.pivot_table(required_fitas, index='CUSTCOD', values='Outstanding (MOP)', aggfunc=np.sum)


    with pd.ExcelWriter(filename) as writer:
        mortgage.to_excel(writer, "mortgage", index=False)
        termloan.to_excel(writer, "termLoan", index=False)
        required_odlimit.to_excel(writer, "odlimit", index=False)
        pivot_odlimit.to_excel(writer, 'pivot_odlimit')
        required_fitas.to_excel(writer, "fitas", index=False)
        pivot_fitas.to_excel(writer, 'pivot_fitas')

    sourcepath = filename
    despath = str(year) + " " + str(month) + "/" + filename

    shutil.move(filename,despath)

