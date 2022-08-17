import os
import pandas as pd
import numpy as np
import shutil


#for i in range(1,66):
    #os.mkdir('2021 0' + str(i))
sa = pd.read_excel('1) BALANCE by CUSTOMER.xlsx', sheet_name=1)
ca = pd.read_excel('1) BALANCE by CUSTOMER.xlsx', sheet_name=2)
fd = pd.read_excel('1) BALANCE by CUSTOMER.xlsx', sheet_name=3)
od = pd.read_excel('1) BALANCE by CUSTOMER.xlsx', sheet_name=4)
start_year = int(input('start year: '))
end_year = int(input('end year: '))
start_month = int(input('start month: '))
end_month = int(input('end month: '))

for year in range(start_year,end_year+1):
    for month in range(start_month, end_month+1):
        required_sa = sa[(sa.Year == year) & (sa.Month == month)]
        required_ca = ca[(ca.Year == year) & (ca.Month == month)]
        required_fd = fd[(fd.Year == year) & (fd.Month == month)]
        required_od = od[(od.Year == year) & (od.Month == month)]


        pivot_sa = pd.pivot_table(required_sa, index='CUSTCOD', values='CurBal/Owner(MOP)',aggfunc=np.sum)
        pivot_ca = pd.pivot_table(required_ca, index='CUSTCOD', values='CurBal/Owner(MOP)',aggfunc=np.sum)
        pivot_fd = pd.pivot_table(required_fd, index='CUSTCOD', values='CurBal/Owner(MOP)',aggfunc=np.sum)
        pivot_od = pd.pivot_table(required_od, index='CUSTCOD', values='CurBal/Owner(MOP)',aggfunc=np.sum)
        filename = str(year) + " " + str(month) + ' balance.xlsx'


        with pd.ExcelWriter(filename) as writer:
            required_sa.to_excel(writer, "sa", index=False)
            pivot_sa.to_excel(writer, 'pivot_sa')
            required_ca.to_excel(writer, "ca", index=False)
            pivot_ca.to_excel(writer, 'pivot_ca')
            required_fd.to_excel(writer, "fd", index=False)
            pivot_fd.to_excel(writer, 'pivot_fd')
            required_od.to_excel(writer, "od", index=False)
            pivot_od.to_excel(writer, 'pivot_od')

        sourcepath = filename
        despath = str(year) + " " + str(month) + "/" + filename

        shutil.move(filename, despath)

