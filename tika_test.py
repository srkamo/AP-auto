import PyPDF2
import xlwt
from xlwt import Workbook


fileName = '.pdf'
pdfFileObj = open(fileName, 'rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

print(pdfReader.numPages)

pageObj = pdfReader.getPage(0)

res = pageObj.extractText()

print(res)

invoice = "Invoice"
ua = "UA"
u2 = "U2"
invoiceDate = "Invoice date"
invoiceLen = 24
orderDate = "Order date"

invoiceNumLength = 6
poNumLen = 16
invoiceIdx = res.find(invoice)
invoiceIdx += len(invoice)

invoiceNum = res[invoiceIdx: invoiceIdx + invoiceNumLength]
print("Invoice: " + str(invoiceNum))
wb = Workbook()

invoiceDateIdx = res.find(invoiceDate)
invoice

uaNumIdx = 0
uaNum = []
while uaNumIdx != -1:
    uaNumIdx = res.find(ua, uaNumIdx)
    if uaNumIdx != -1:
        uaNum.append(res[uaNumIdx: uaNumIdx + poNumLen])
        uaNumIdx += poNumLen


u2NumIdx = 0
u2Num = []
while u2NumIdx != -1:
    u2NumIdx = res.find(u2, u2NumIdx)
    if u2NumIdx != -1:
        u2Num.append(res[u2NumIdx: u2NumIdx + poNumLen])
        u2NumIdx += poNumLen



print("PO#: " + str(u2Num))
sheet1 = wb.add_sheet('Sheet 1')
#(row, col, 'content')
#sheet1.write(1, 0, uaNum)
sheet1.write(1, 1, invoiceNum)
sheet1.write(2, 0, u2Num)
sheet1.write(2, 1, invoiceNum)
sheet1.write(3, 0, '3, 0')
sheet1.write(0, 1, 'yo')
sheet1.write(0, 2, '0, 2')
sheet1.write(0, 3, '0, 3')

wb.save('test.xls')
