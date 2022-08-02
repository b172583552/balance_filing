import os
import pandas as pd

def extractText(textfile, code, year, month):
    requiredlist = []
    accountlist = []
    outstandinglist=[]
    matchstr = "Outstanding:"
    outstanding = 0
    currency = ['HKD', 'RMB', 'USD', 'GBP']
    currency_list = pd.read_excel('AMCM Monthend Closing Rates 2017-2021.xlsx', header=0)
    currency_list = currency_list[currency_list['Date'].str.contains(year, na=False)]
    currency_list = currency_list[currency_list['Date'].str.contains(month, na=False)]
    currency_list.columns = ['Date', 'HKD', 'RMB', 'USD', 'GBP']
    with open(textfile) as fo:
        for rec in fo:
            if rec[0:2] == code:
                requiredlist.append(outstanding)
                outstanding = 0
                output = ""
                for i in rec:
                    output += i
                    if i == " ":
                        requiredlist.append(output)
                        break
            for k in currency:
                if k in rec:
                    curr_currency = currency_list.iloc[0][k]
            if matchstr in rec:
                temp = ""
                for char in reversed(rec[0:rec.find(matchstr)]):
                    temp += char
                    if char == " ":
                        break
                temp = temp[::-1]
                temp = temp.replace(",", "")
                temp = temp.strip("\n")
                temp = temp.strip()
                outstanding += (float(temp) * curr_currency)
        requiredlist.append(outstanding)
    del requiredlist[0]
    for index, obj in enumerate(requiredlist):
        if index % 2 == 0:
            accountlist.append(obj)
        else:
            outstandinglist.append(obj)
    return accountlist, outstandinglist

if __name__ == '__main__':
    accountlist, outstandinglist = extractText('2021 6 bond.txt', 'BM', '2021', '6')
    print(accountlist)
    print(outstandinglist)
