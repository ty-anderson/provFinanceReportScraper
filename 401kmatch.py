import xlwings as xw
import pandas as pd
import math

wb = xw.Book(r'C:\Users\tyler.anderson\Desktop\PAYROLL & CENSUS DATA - Bi-Weekly - 2020.xlsm')

a = 'L'
x = 1
y = 1
building_list = []
value_list = []
account_list = ['6110.26', '6200.26', '6300.26', '6400.26', '6500.26', '6600.26', '6700.26',
                '6800.26', '6900.261', '6900.261', '6900.261', '6900.261', '6900.261',
                '8200.26', '8200.26', '8200.26', '8250.26', '8250.26', '8280.26', '8280.26'
                ]
gl_list = []
description_list = []
date_list = []
calc_list = []
empty_list = []
fullamt_list = []

for sheet in wb.sheets:
    building_list.append(sheet.name)
    fullamt = sheet.range(a + '702').value
    try:
        amt = round(fullamt * .22, 2)
        calc_list.append('Calculated')
    except:
        amt = fullamt
        calc_list.append('not adjusted')
    value_list.append(amt)
    description_list.append('401k Match Accrual')
    date_list.append(sheet.range(a + '8').value)
    fullamt_list.append(fullamt)
    empty_list.append(' ')
    for i in range(704, 723):
        building_list.append(sheet.name)
        fullamt = sheet.range(a + str(i)).value
        try:
            amt = round(fullamt * .22, 2)
            calc_list.append('Calculated')
        except:
            amt = fullamt
            calc_list.append('not adjusted')
        value_list.append(amt)
        description_list.append('401k Match Accrual')
        date_list.append(sheet.range(a + '8').value)
        fullamt_list.append(fullamt)
        empty_list.append(' ')
    for i in account_list:
        gl_list.append(i)


master_list = zip(empty_list, gl_list, description_list, value_list, empty_list, date_list, empty_list, empty_list, empty_list, empty_list, building_list, calc_list,fullamt_list)
df = pd.DataFrame(master_list)
count = max(df.index.to_list())
print(count)
je = xw.Book()
je.sheets[0].range('A1').value = df
je.sheets[0].range('A1').value = 'FACILITY'
je.sheets[0].range('B1').value = 'REFERENCE_DESCRIPTION'
je.sheets[0].range('C1').value = 'ACCOUNT'
je.sheets[0].range('D1').value = 'DESCRIPTION'
je.sheets[0].range('E1').value = 'DEBIT'
je.sheets[0].range('F1').value = 'CREDIT'
je.sheets[0].range('G1').value = 'EFFECTIVE_DATE'
je.sheets[0].range('H1').value = 'FISCAL_YEAR'
je.sheets[0].range('I1').value = 'FISCAL_PERIOD'
je.sheets[0].range('J1').value = ''
je.sheets[0].range('K1').value = ''
je.sheets[0].range('L1').value = 'Facility'
je.sheets[0].range('M1').value = ''
je.sheets[0].range('H2').value = '=YEAR(G2)'
je.sheets[0].range('I2').value = '=NUMBERVALUE(TEXT(G2,"MM"))'



