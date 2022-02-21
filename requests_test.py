import requests

f = open("info.txt", "r")
u = f.readline().split(',')
f.close()

s = requests.Session()
payload = {'username': u[0], 'password': u[1]}
url = "https://login.pointclickcare.com/home/userLogin.xhtml"
r = s.post(url, data=payload)

r = s.get("https://www30.pointclickcare.com/glap/reports/rp_customglreports.jsp?ESOLrepId=255")

payload = {'ESOLclientid': '',
           'client_id': '-1',
           'EnablePageSetup': '',
           'ESOLrepId': '255',
           'ESOLrepType': 'PB',
           'ESOLfiscalYrEnd': '12',
           'ESOLnoOfPeriods': '12',
           'ESOLRepFileName': 'PGHC Reporting Income Statement (Excel Rpt Export)',
           'ESOLminiToken': 'gq4wfneil24',
           'facType': 'USAR',
           'ESOLperiod': 1,
           'ESOLmonth': 1,
           'ESOLstartper': 1,
           'ESOLendper': 12,
           'ESOLstartyr': 2021,
           'ESOLendyr': 2021,
           'ESOLstartperiod': 1,
           'ESOLyear': 2021,
           'ESOLendperiod': 12,
           'ESOLshowShading': 'Y',

           }
r = s.post("https://www30.pointclickcare.com/glap/reports/rp_customglreports.jsp?ESOLrepId=255", data=payload)
print(r)
