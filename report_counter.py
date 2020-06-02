from os import listdir, environ
from os.path import isfile, join
import datetime
from collections import Counter

"""
This will count how many reports have been run for each facility, informing you of ones that are missing.
If it is not the same as the reporting month number (by May there should be 5 reports for each)
then it will create a report on the desktop to inform you so you can run the missing months.
"""
reportmonth = datetime.date.today().month - 1
userpath = environ['USERPROFILE']


def to_text(message):
    s = str(str(message) + "\n")
    with open(userpath + '\\Desktop\\Month End Reports.txt','a') as file:
        file.write(s)
        file.close()


mypath = [r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries',
          r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation']

for path in mypath:
    check_list = []
    full_list = []
    for f in listdir(path):
        year = f[:4]
        month = f[5:7]
        fsplit = f[8:]
        check_list.append(fsplit)
        full_list.append(f)

    c = Counter(check_list)

    for item in c:
        if c[item] != reportmonth:
            to_text(item + " " + str(c[item]))

to_text('Done')
