import os
import glob
import pandas as pd
import numpy as np
from PyPDF2 import PdfReader
from data_extractor import extractText
import shutil
import re
#for i in range(1,66):
    #os.mkdir('2021 0' + str(i))

def pdftotext(file_read, file_write):
    reader = PdfReader(file_read)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    with open(file_write, "w") as file:
        file.write(text)

def bondFile(year, month):
    bond = pd.DataFrame()
    special_month=[10,11,12]
    if month not in special_month:
        path = 'C:/Users/tomwong/Desktop/Bonds/' + str(year) + ' Reports/??-0' + str(month) + "*"
    else:
        path = 'C:/Users/tomwong/Desktop/Bonds/' + str(year) + ' Reports/??-' + str(month) + "*"
    for file in glob.glob(path):
        print(file)
        file_read = file
        file_write = str(year) + " " + str(month) + " bond.txt"
        pdftotext(file_read, file_write)
    accountlist, outstandinglist = extractText(file_write, 'BM', str(year), str(month))
    bond['ACCNO'] = accountlist
    bond['ACCNO'] = bond['ACCNO'].str.strip()
    bond['Values'] = outstandinglist
    bond = pd.merge(bond,mapping.iloc[:, 2:5], on=['ACCNO'], how='left')
    bond['CCY'] = 'MOP'
    bond['AccType'] = 'bond'
    bond['Year'] = year
    bond['Month'] = month
    bond['bookrate'] = 1
    bond['outstanding per owner'] = bond.apply(lambda row: row['Values']/row['Account Owners'], axis=1)
    bond = bond[['CUSTCOD', 'AccType', 'ACCNO', 'CCY', 'Year', 'Month', 'Values','Account Owners', 'bookrate', 'outstanding per owner']]
    print(bond)
    return bond

def fundFile(year,month):
    fund = pd.DataFrame()
    special_month = [10, 11, 12]
    if month not in special_month:
        path = 'C:/Users/tomwong/Desktop/Funds/' + str(year) + ' Reports/??-0' + str(month) + "*"
    else:
        path = 'C:/Users/tomwong/Desktop/Funds/'+str(year)+' Reports/??-' + str(month) + "*"
    for file in glob.glob(path):
        print(file)
        file_read = file
        file_write = str(year) + " " + str(month) + " fund.txt"
        pdftotext(file_read,file_write)
    accountlist, outstandinglist = extractText(file_write, 'FM', str(year), str(month))
    fund['ACCNO'] = accountlist
    fund['ACCNO'] = fund['ACCNO'].str.strip()
    fund['Values'] = outstandinglist
    fund = pd.merge(fund,mapping.iloc[:, 2:5], on=['ACCNO'], how='left')
    fund['CCY'] = 'MOP'
    fund['AccType'] = 'fund'
    fund['Year'] = year
    fund['Month'] = month
    fund['bookrate'] = 1
    fund['outstanding per owner'] = fund.apply(lambda row: row['Values']/row['Account Owners'], axis=1)
    fund = fund[['CUSTCOD', 'AccType', 'ACCNO', 'CCY', 'Year', 'Month', 'Values','Account Owners', 'bookrate', 'outstanding per owner']]
    print(fund)
    return fund




def securitiesFile(year,month):
    special_month = [10, 11, 12]
    if year == 2017:
        if month not in special_month:
            path = 'C:/Users/tomwong/Desktop/Securities/' + str(year) + "/" + "*20170" + str(month) + "*"
        else:
            path = 'C:/Users/tomwong/Desktop/Securities/' + str(year) + "/" + "*2017" + str(month) + "*"

    else:
        if month not in special_month:
            path = 'C:/Users/tomwong/Desktop/Securities/' + str(year) + "/" + "0" + str(month) + "*"
        else:
            path = 'C:/Users/tomwong/Desktop/Securities/' + str(year) + "/" + str(month) + "*"


    for file in glob.glob(path):
        print(file)
        securities = pd.read_excel(file, header=None)
        if year == 2018 and month == 5:
            print('special case')
            securities = securities[securities.iloc[:, 0].str.contains('C{1}', na=False)]
            securities['outstanding'] = securities.ffill(axis=1).iloc[:, -1]
            extractSecurities = securities.iloc[:, [1, -1]]
            extractSecurities.dropna(inplace=True)
            extractSecurities.columns = ['ACCNO', 'Values']
            extractSecurities['Values'] = extractSecurities['Values'].apply(lambda x: str(x).split()[-1].replace(",", ''))
            extractSecurities['Values'] = extractSecurities['Values'].apply(lambda x: np.nan if re.search(r'[A-Za-z\(\):]', x) else float(x))
        elif file.endswith('.xlsx'):
            print('file end with xlsx')
            securities = securities[securities.iloc[:, 0].str.contains('Sub-Total', na=False)]
            securities['outstanding'] = securities.ffill(axis=1).iloc[:, -1]
            extractSecurities = securities.iloc[:, [0, -1]]
            extractSecurities.columns = ['ACCNO', 'Values']
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.extract(r'(\(.*?\))')
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.replace(')', '')
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.replace('(', '')
            extractSecurities['Values'] = extractSecurities['Values'].apply(lambda x: str(x).split()[-1].replace(",", ''))
            extractSecurities['Values'] = extractSecurities['Values'].apply(lambda x: np.nan if re.search(r'[A-Za-z\(\):]', x) else float(x))
        else:
            securities = securities[securities.iloc[:, 0].str.contains('Sub-Total', na=False)]
            extractSecurities = securities.iloc[:, [0, 21]]
            extractSecurities.columns = ['ACCNO', 'Values']
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.extract(r'(\(.*?\))')
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.replace(')', '')
            extractSecurities['ACCNO'] = extractSecurities['ACCNO'].str.replace('(', '')

        extractSecurities = pd.merge(extractSecurities, mapping.iloc[:, 2:5], on=['ACCNO'], how='left')
        extractSecurities['CCY'] = 'MOP'
        extractSecurities['AccType'] = 'securities'
        extractSecurities['Year'] = year
        extractSecurities['Month'] = month
        extractSecurities['bookrate'] = 1.03
        extractSecurities['outstanding per owner'] = extractSecurities.apply(lambda row: row['Values']/row['Account Owners'] * row['bookrate'], axis=1)
        extractSecurities = extractSecurities[['CUSTCOD', 'AccType', 'ACCNO', 'CCY', 'Year', 'Month', 'Values', 'Account Owners', 'bookrate','outstanding per owner']]
        print(extractSecurities)
    return extractSecurities

def exportExcel(bond, fund, securities, pivot_securities, pivot_bond, pivot_fund, filename):
    with pd.ExcelWriter(filename) as writer:
        securities.to_excel(writer, 'securities', index=False)
        pivot_securities.to_excel(writer, 'pivot_securities')
        fund.to_excel(writer, 'fund', index=False)
        pivot_fund.to_excel(writer, 'pivot_fund')
        bond.to_excel(writer, 'bond', index=False)
        pivot_bond.to_excel(writer, 'pivot_bond')
    print('export successfully!')


if __name__ == "__main__":
    start_year = int(input('start year: '))
    end_year = int(input('end year: '))
    start_month = int(input('start month: '))
    end_month = int(input('end month: '))
    mapping = pd.read_excel('3) Investment by CUSTOMER.xlsx', sheet_name=3)
    bookrate = pd.read_excel('3) Investment by CUSTOMER.xlsx', sheet_name=4)
    for year in range(start_year,end_year+1):
        for month in range(start_month,end_month+1):
            filename = str(year) + " " + str(month) + ' investment.xlsx'
            securities = securitiesFile(year, month)
            bond = bondFile(year, month)
            fund = fundFile(year, month)


            pivot_securities = pd.pivot_table(securities, index='CUSTCOD', values='outstanding per owner', aggfunc=np.sum)
            pivot_bond = pd.pivot_table(bond, index='CUSTCOD', values='outstanding per owner', aggfunc=np.sum)
            pivot_fund = pd.pivot_table(fund, index='CUSTCOD', values='outstanding per owner', aggfunc=np.sum)

            exportExcel(bond, fund, securities, pivot_securities, pivot_bond, pivot_fund, filename)

            sourcepath = filename
            despath = str(year) + " " + str(month) + "/" + filename
            shutil.move(filename,despath)