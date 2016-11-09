# coding=utf-8
# Sephora order tracker
# create by Haokun Lin
import re
import urllib
import urllib2
import cookielib
import xlrd
import xlwt
from bs4 import BeautifulSoup


def getAllOrders():
    orderFile = xlrd.open_workbook('allorders.xlsx',encoding_override='utf-8')
    orderTable = orderFile.sheets()[0]
    itemList=[]
    orderList=[]
    allOrderList=[]
    for rownum in range(1,orderTable.nrows):
        tempPersonName = orderTable.cell(rownum,2).value
        if tempPersonName!='':
            if len(orderList)!=0:
                orderList.append(itemList)
                allOrderList.append(orderList)
            orderList = []
            itemList = []
            personName = tempPersonName
            orderList.append(personName)
        itemCode=str(orderTable.row_values(rownum)[3])
        if itemCode!='':
            itemCode=int(orderTable.row_values(rownum)[3])
            itemQty = int(orderTable.row_values(rownum)[4])
            itemList.append({'ID': itemCode,'Qty':itemQty})
    orderList.append(itemList)
    allOrderList.append(orderList)
    return allOrderList


accountData = xlrd.open_workbook('sephoraaccount.xlsx')
accountTable = accountData.sheets()[0]

resultFile = xlwt.Workbook()
resultTable = resultFile.add_sheet('orderDetail')

resultTable.write(0, 0, 'Order Number')
resultTable.write(0, 1, 'Tracking Number')
resultTable.write(0, 2, 'Item Number')
resultTable.write(0, 3, 'Price')
resultTable.write(0, 4, 'Qty')
resultTable.write(0, 5, 'Amount')
resultTable.write(0, 6, 'Account')

rowNumber = 1

for rownum in range(accountTable.nrows):
    userName = accountTable.cell(rownum, 0).value
    password = accountTable.cell(rownum, 1).value
    print "Account: "+userName
    resultTable.write(rowNumber, 6, userName)


    filename = 'cookie.txt'
    cookie = cookielib.MozillaCookieJar(filename)
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie))
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    postdata = urllib.urlencode({
        'password': password,
        'user_name':userName
    })

    detailUrl='https://www.sephora.com/profile/orders/orderDetail.jsp?orderId='
    loginUrl = 'https://www.sephora.com/rest/user/login'
    result = opener.open(loginUrl, postdata)
    cookie.save(ignore_discard=True, ignore_expires=True)
    gradeUrl = 'https://www.sephora.com/profile/orders/orderHistory.jsp'
    result = opener.open(gradeUrl)
    soup = BeautifulSoup(result,"html.parser")
    for order in soup.find_all('table',class_="Table u-dt--fixed"):
        try:
            orderDate=next(order.parent.parent.find('td', text=re.compile(r'November.*')).stripped_strings)
        except:
            orderDate='NONE'
        if orderDate!='NONE':
            orderId=next(order.find('a',class_='u-hoverRed u-underline').stripped_strings)
            tracking=''
            try:
                tracking=next(order.find('a', class_='js-pop-window u-hoverRed u-underline').stripped_strings)
            except:
                tracking='NONE'
            print 'Order Number: '+orderId+' Tracking number: '+tracking

            if tracking!='NONE':

                detailResult = opener.open(detailUrl + orderId)
                detailRead=detailResult.read()

                addressReg = r'68512'
                addressPattern = re.compile(addressReg,re.I)
                addressFind = addressPattern.findall(detailRead)
                if(len(addressFind)!=0):
                    resultTable.write(rowNumber, 0, orderId)
                    resultTable.write(rowNumber, 1, tracking)
                    detailReg = r'"sku_number":"(\d*)"'
                    detailPattern = re.compile(detailReg)
                    detailFind = detailPattern.findall(detailRead)
                    print '----------------------------------------'
                    detailResult = opener.open(detailUrl + orderId)
                    detailSoup = BeautifulSoup(detailResult, "html.parser")
                    i=0
                    for item in detailSoup.find_all('tr',attrs={"data-at": "order_item"}):
                        itemNumber = detailFind[i]
                        amount=next(item.find('td',attrs={"data-at": "order_item_amt"}).stripped_strings)
                        if amount!='$0.00':
                            qty=next(item.find('td',attrs={"data-at": "order_item_qty"}).stripped_strings)
                            price=next(item.find('td',attrs={"data-at": "order_item_price"}).stripped_strings)
                            print '|Item Number: '+itemNumber+' Price: '+price+' Qty: '+qty+' Amount: '+amount
                            resultTable.write(rowNumber, 2, itemNumber)
                            resultTable.write(rowNumber, 3, price)
                            resultTable.write(rowNumber, 4, qty)
                            resultTable.write(rowNumber, 5, amount)
                            rowNumber=rowNumber+1
                        i=i+1
                    rowNumber = rowNumber + 1
            print '****************************************'

resultFile.save('orderDetails.xls')