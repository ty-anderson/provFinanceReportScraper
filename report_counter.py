from os import listdir, environ
import datetime
import xlwings as xw
import pandas as pd

"""
This will count how many reports have been run for each facility, informing you of ones that are missing.
If it is not the same as the reporting month number (by May there should be 5 reports for each)
then it will create a report on the desktop to inform you so you can run the missing months.
"""
reportmonth = datetime.date.today().month - 1
userpath = environ['USERPROFILE']

wb_ref = r"P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx"
wb = pd.read_excel(wb_ref, sheet_name='Automation', usecols=['Common Name'])

wb_list = wb['Common Name'].to_list()

print(wb_list)

mypath = [r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AP Aging',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation']

wb = xw.Book()

for path in mypath:
    if path == mypath[0]:
        tabname = "AP Aging"
    if path == mypath[1]:
        tabname = "AR Aging"
    if path == mypath[2]:
        tabname = "AR Rollforward"
    if path == mypath[3]:
        tabname = "Cash Receipts"
    if path == mypath[4]:
        tabname = "Census"
    if path == mypath[5]:
        tabname = "Journal Entries"
    if path == mypath[6]:
        tabname = "Revenue Recon"
    wb.sheets.add(tabname)
    sht = wb.sheets[tabname]
    x = 1
    y = 1
    col_header = '2020 01'
    for f in listdir(path):
        s = f.split()
        if len(s) > 1:
            year = s[0]
            month = s[1]
            if y > 1:
                try:
                    if sht.range(x, y - 1).value[8:] != f[8:]:
                        sht.api.rows(x).insert
                        sht.range(x, y).color = (255, 94, 94)
                except:
                    pass
            if col_header != str(year + ' ' + month):
                y += 1
                x = 1
                col_header = str(year + ' ' + month)
            sht.range(x, y).value = f
            x += 1
    wb.sheets[tabname].autofit('c')

wb.sheets['Sheet1'].delete()
