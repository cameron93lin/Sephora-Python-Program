import xlrd
import xlwt
import urllib2
import re
import requests
import json

data = xlrd.open_workbook('sephoragc.xlsx')
table = data.sheets()[0]
nrows = table.nrows


file = xlwt.Workbook()
table2 = file.add_sheet('giftcard')

for rownum in range(table.nrows):
    url = 'https://www.sephora.com/rest/giftcards?t=1478546517576&gc_number=' + table.row_values(rownum)[0] + '&gc_pin=' + table.row_values(rownum)[1] + '&country_switch=us'
    table2.write(rownum, 0,table.row_values(rownum)[0])
    table2.write(rownum, 1, table.row_values(rownum)[1])
    detailRequest = urllib2.Request(url,headers={'User-Agent':'Mozilla/5.0'})
    detailResponse = urllib2.urlopen(detailRequest)
    data = json.load(detailResponse)
    table2.write(rownum, 2, data[u'balance'])
    print 'Card: '+table.row_values(rownum)[0]
    print 'Balance: '+str(data[u'balance'])

file.save('giftcards.xls')
