import calendar
import shutil
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import chromedriver_autoinstaller
import json
import time
import datetime
import os
import xlwings as xw
import pyautogui
import win32com
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QGridLayout, QWidget, QCheckBox, \
    QPushButton, QVBoxLayout, QFrame, QFormLayout, QLineEdit, QTextEdit
from PyQt5.QtCore import QSize
import sys
import pyperclip
from threading import Thread, Event
from queue import Queue

"""
This app is a web scraper for pulling reports from PCC
1. Month end reports for PACS close
2. Income statements for PACS financial package to be converted to the 96-line income statement
3. Census report of kindred buildings census
4. Intercompany reports (balance sheet-system & GL transactions for 1340.000)

To update when new buildings are acquired "P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx"

The process currently checks every 59 seconds to see if today is the 15th at 8:00pm to run month end reports. This must
be running at that time for it to work which means this needs to constantly run in the background.  Would be better on 
a server with access to the shared drive (probably not going to happen so run manually, currently commented out).

Files (Chromedriver and pccwebscraping) are hosted on shared drive but moved to local folder PCC HUB in users documents 
folder to run.
"""

# clear the gen_py folder that is causing issues with the xlsx conversion with win32com
try:
    shutil.rmtree(win32com.__gen_path__[:-4])
except:
    pass

global newpathtext
global PCC
global q
global event

q = Queue()
event = Event()

# collect user info
username = os.environ['USERNAME']
userpath = os.environ['USERPROFILE']

"""Setup date info"""
today = datetime.date.today()
current_year = today.year
prev_month_num = today.month - 1
if prev_month_num == 0:
    prev_month_num = 12
    report_year = today.year - 1
if len(str(prev_month_num)) == 1:
    prev_month_num_str = str("0" + str(prev_month_num))
else:
    prev_month_num_str = str(prev_month_num)
prev_month_abbr = calendar.month_abbr[prev_month_num]
prev_month_word = calendar.month_name[prev_month_num]
if prev_month_num != 12:
    report_year = today.year

"""Get paths to map out how data flows if not connected to the VPN"""
try:
    # faclistpath = 'P:\\PACS\\Finance\\General Info\\Finance Misc\\Facility List.xlsx'
    faclistpath = "P:\\PACS\\Finance\\Automation\\PCC Reporting\\pcc webscraping.xlsx"
    try:
        os.mkdir(userpath + '\\Documents\\PCC HUB\\')  # make directory for backup in documents folder
        shutil.copyfile(faclistpath, userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx')  # make backup file
    except FileExistsError:
        shutil.copyfile(faclistpath,
                        userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx')  # if folder exists just copy
except FileNotFoundError:  # if VPN is not connected use the one last saved
    faclistpath = userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx'

facility_df = pd.read_excel(faclistpath, sheet_name='Automation', index_col=0)

"""Create All lists and dictionaries"""
facilityindex = facility_df.index.to_list()
accountants = facility_df['Accountant'].to_list()
fac_number = facility_df['Business Unit'].to_list()
pcc_name = facility_df['PCC Name'].to_list()
facilities = dict(zip(facilityindex, zip(accountants, fac_number, pcc_name)))
reports_list = ['AP Aging',
                'AR Aging',
                'AR Rollforward',
                'Cash Receipts',
                'Detailed Census',
                'Journal Entries',
                'Revenue Reconciliation']


def update_date(monthinput='', yearinput=''):
    """If date info needs to change"""
    global prev_month_num_str
    global prev_month_word
    global prev_month_num
    global prev_month_abbr
    global report_year
    # collect date info
    try:
        prev_month_num = int(monthinput)
    except ValueError:
        prev_month_num = prev_month_num
    if len(str(prev_month_num)) == 1:
        prev_month_num_str = str("0" + str(prev_month_num))
    else:
        prev_month_num_str = str(prev_month_num)
    prev_month_abbr = calendar.month_abbr[prev_month_num]
    prev_month_word = calendar.month_name[prev_month_num]
    try:
        report_year = int(yearinput)
    except ValueError:
        report_year = report_year
    print('Reporting date is ' + prev_month_abbr + ' ' + str(report_year))


def check_reports():
    counter = 0
    wb = pd.read_excel(r"P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx", sheet_name='Automation',
                       usecols=['Common Name', 'Business Unit'])
    reports_path = [r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AP Aging',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation']
    report_names = ['AP Aging.xlsx', 'AR Aging.xlsx', 'AR Rollforward.xlsx', 'Cash Receipts.pdf',
                    'Census.pdf', 'Journal Entries.pdf', 'Revenue Reconciliation.pdf']
    # q.put("Searching monthly reports")
    print("Searching monthly reports")
    i = 0
    for path in reports_path:
        for building in wb['Common Name']:
            file_name = path + '\\' + str(report_year) + ' ' + str(prev_month_num_str) + ' ' + building + ' ' + \
                        report_names[i]
            if not os.path.exists(file_name):
                counter = counter + 1
                # q.put(file_name + ' missing.  Downloading now')
                print(file_name + ' missing.  Downloading now')
                rpt = report_names[i].split('.')
                rpt = [rpt[0]]
                download_reports(building, rpt)
            else:
                size = os.path.getsize(file_name)
                kb = size / 1024
                if kb < 10:
                    # q.put(file_name + " might be empty.  Please check.")
                    print(file_name + " might be empty.  Please check.")
            # q.put(f"Completed {building} {file_name}")
            # print(f"Completed {building} {file_name}")
        i += 1
    if counter == 0:
        # q.put("Reports have all been downloaded")
        # event.set()
        print("Reports have all been downloaded")


def deleteDownloads():
    """Deletes everything in downloads folder"""
    filelist = glob.glob(userpath + '\\Downloads\\*')
    try:
        for f in filelist:
            os.remove(f)
    except:
        pass


def renameDownloadedFile(newfilename, dirpath=''):
    """Renames most recent file in downloads folder and moves it to dirpath"""
    global newpathtext
    try:
        newptext = newpathtext
    except NameError:
        newptext = '\\'
    time.sleep(2)
    try:
        listoffiles = glob.glob(userpath + '\\Downloads\\*')  # get a list of filesp
        latestfile = max(listoffiles, key=os.path.getctime)  # find the latest file
        extention = os.path.splitext(latestfile)[1]  # get the extension of the latest file
        if dirpath == '':
            os.rename(latestfile, userpath + '\\Downloads\\' + newfilename)
        else:
            if newptext != '\\':
                destfile = os.path.join(newptext, newfilename + extention)
            else:
                destfile = os.path.join(dirpath, newfilename + extention)
            try:
                shutil.move(latestfile, destfile)  # try to save file to original folder (if error with VPN)
            except:  # BACKUP LOCATION IF VPN GOES DOWN
                try:  # make new folder if doesn't exist
                    os.mkdir(userpath + '\\Desktop\\temp reporting\\')  # temp file on desktop
                    newptext = userpath + '\\Desktop\\temp reporting\\'  # temp file on desktop
                    destfile = os.path.join(newptext, newfilename + extention)  # form save file path
                except FileExistsError:  # if folder does exist then just save
                    newptext = userpath + '\\Desktop\\temp reporting\\'
                    destfile = os.path.join(newptext, newfilename + extention)
                shutil.move(latestfile, destfile)  # MOVE AND RENAME
            print('Moved to: ' + destfile)  # END BACKUP LOCATION
    except:
        print("Issue renaming/moving to " + str(dirpath))
        print(newfilename + " is in Downloads folder")


def convert_to_xlsx():
    """Opens non-xlsx file and saves as xlsx"""
    listoffiles = glob.glob(userpath + '\\Downloads\\*')  # get a list of files
    latestfile = max(listoffiles, key=os.path.getctime)  # find the latest file
    extention = os.path.splitext(latestfile)[1]  # get the extension of the latest file
    excel = win32com.client.dynamic.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(latestfile)
    wb.SaveAs(latestfile + "x", FileFormat=51)
    wb.Close()
    excel.Application.Quit()


def check_if_downloaded(facility, report):
    time.sleep(3)
    if report == "Cash Receipts":
        report_name = "Cash Receipts.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts'
    elif report == "AP Aging":
        report_name = "AP Aging.xlsx"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AP Aging'
    elif report == "AR Aging":
        report_name = "AR Aging.xlsx"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging'
    elif report == "AR Rollforward":
        report_name = "AR Rollforward.xlsx"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward'
    elif report == "Census":
        report_name = "Census.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census'
    elif report == "Journal Entries":
        report_name = "Journal Entries.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries'
    elif report == "Revenue Reconciliation":
        report_name = "Revenue Reconciliation.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation'
    else:
        # report_name = "Issue identifying report"
        # report_path = "Issue identifying report"
        pass
    file_name = report_path + '\\' + str(report_year) + ' ' + str(
        prev_month_num_str) + ' ' + facility + ' ' + report_name
    if not os.path.exists(file_name):
        print(file_name + ' missing')


def find_updated_driver():
    """Pulls latest driver from shared drive and addes to user's documents folder"""
    folder = 'P:\\PACS\\Finance\\Automation\\Chromedrivers\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        try:
            print('Updating chromedriver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + max(file_list) + '.exe')
            print('chromedriver updated to version ' + max(file_list))
        except:
            print("Couldn't update chromedriver automatically")
        return max(file_list)
    else:
        print('Could not find P:\\PACS\\Finance\\Automation\\Chromedrivers\\')


def find_current_driver():
    """Uses driver that is stored on your computer"""
    folder = os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        return max(file_list)


def startPCC():
    """Start new instance of class PCC"""
    global PCC
    PCC = LoginPCC()


def download_reports(facilitylist=facilityindex, reportlist=reports_list):
    """Download month end close reports"""
    global PCC
    deleteDownloads()
    if not facilitylist:
        facilitylist = facilityindex
    if reportlist:
        try:
            PCC  # check if an instance already exists
        except:  # if not
            startPCC()  # create one
        for facname in facilities:  # LOOP BUILDING LIST
            if facname in facilitylist:  # IS BUILDING CHECHED
                bu = str(facilities[facname][1])  # GET BU
                if len(bu) < 2:
                    bu = str(0) + bu
                if PCC.buildingSelect(bu):
                    time.sleep(1)
                    for report in reportlist:
                        if report == 'AP Aging':
                            PCC.ap_aging(facname)
                        if report == 'AR Aging':            # USES MGMT CONSOLE
                            bu = facilities[facname][1]     # TO SELECT BUILDING IN AR REPORT
                            PCC.ar_aging(facname, bu)
                            PCC.buildingSelect(str(bu))
                        if report == 'AR Rollforward':
                            PCC.ar_rollforward(facname)
                        if report == 'Cash Receipts':
                            PCC.cash_receipts(facname)
                        if report == 'Census':
                            PCC.census(facname)
                        if report == 'Journal Entries':
                            PCC.journal_entries(facname)
                        if report == 'Revenue Reconciliation':
                            PCC.revenuerec(facname)
                        check_if_downloaded(facname, report)
        print('Reports downloaded')
    else:
        print('No reports selected.')


class LoginPCC:
    def __init__(self):
        """Create instance, login to PCC"""
        try:
            chrome_options = webdriver.ChromeOptions()
            settings = {
                "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }
            prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                     "plugins.always_open_pdf_externally": True}
            chrome_options.add_experimental_option('prefs', prefs)
            chrome_options.add_argument('--kiosk-printing')
            chromedriver_autoinstaller.install()
            self.driver = webdriver.Chrome(options=chrome_options)

            try:
                self.driver.get('https://login.pointclickcare.com/home/userLogin.xhtml')
                time.sleep(3)
                f = open("info.txt", "r")
                u = f.readline().split(',')
                f.close()
                try:
                    username = self.driver.find_element(By.ID, 'username')
                    username.send_keys(u[0])
                    password = self.driver.find_element(By.ID, 'password')
                    password.send_keys(u[1])
                    self.driver.find_element(By.ID, 'login-button').click()
                    time.sleep(3)
                except:
                    usernamex = self.driver.find_element(By.ID, 'id-un')
                    usernamex.send_keys(u[0])
                    passwordx = self.driver.find_element(By.ID, 'password')
                    passwordx.send_keys(u[1])
                    self.driver.find_element(By.ID, 'id-submit').click()
            except:
                print("There is an issue with the chrome driver")
        except:
            print('There was an issue initiating chromedriver')

    def teardown_method(self):
        """Exit browser"""
        self.driver.quit()

    def close_all_windows(self, firstwindow):
        """Close all windows except for original window"""
        original_window = firstwindow
        all_windows = self.driver.window_handles
        for window in all_windows:
            if window != original_window:
                self.driver.switch_to.window(window)
                self.driver.close()
        self.driver.switch_to.window(firstwindow)

    def buildingSelect(self, bu):
        """Select the building using business unit"""
        try:
            time.sleep(5)
            self.driver.find_element(By.XPATH, '//button[normalize-space()="close"]').click()
        except:
            pass
        try:
            current_fac = self.driver.find_element(By.NAME, "current_fac_id").get_attribute("value")
            if str(current_fac) != bu:
                try:
                    self.driver.find_element(By.ID, "pccFacLink").click()
                    time.sleep(1)
                    building_list = self.driver.find_element(By.ID, "optionList")
                    options_split = building_list.text.splitlines()
                    for option in options_split:
                        bu_pull = option.replace(" ", "")
                        bu_pull = bu_pull[-5:].split("-")
                        bu_val = bu_pull[1]
                        if bu_val == bu:
                            print(option)
                            building_list.find_element(By.PARTIAL_LINK_TEXT, option).click()
                            return True
                except:
                    print("Could not locate " + bu + " in PCC")
                    return False
            else:
                return True
        except:
            print("Could not find the building dropdown menu")
            return False

    def ap_aging(self, facname):
        """Download AP aging report (paste to Excel)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/glap/reports/rp_aptrialbalance.xhtml")
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "tr:nth-child(3) label:nth-child(3) > input").click()
            time.sleep(2)
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
                pyperclip.copy("Fiscal period has not been setup for this entity.  No transactions recorded.")
                wb = xw.Book()  # new workbook
                app = xw.apps.active
                time.sleep(2)
                wb.activate(steal_focus=True)  # focus the new instance
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'v')  # paste
                wb.sheets[0].range("A1:P20").color = (102, 153, 255)
                wb.sheets[0].range("A1:P20").api.Font.Bold = True
                time.sleep(2)  # wait to load
                wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AP Aging\\" +
                        str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                app.quit()
                print(facname + ' AP aging saved to shared drive')
            except:
                pass
            self.driver.find_element(By.NAME, "ESOLmonth").click()
            time.sleep(1)
            dropdown = self.driver.find_element(By.NAME, "ESOLmonth")
            dropdown.find_element(By.CSS_SELECTOR, "td:nth-child(3) > select:nth-child(1) > option:nth-child(" + str(
                prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonth").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyear"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.NAME, "ESOLreportOutputType").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLreportOutputType")
            dropdown.find_element(By.XPATH, "//option[. = 'HTML']").click()
            self.driver.find_element(By.NAME, "ESOLreportOutputType").click()
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(2)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            pyperclip.copy('')
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'a')  # highlight the entire page
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # COPY ALL CONTENT
            time.sleep(2)
            self.close_all_windows(window_before)
            wb = xw.Book()  # new workbook
            app = xw.apps.active
            win_wb = wb.api
            module = win_wb.VBProject.VBComponents.Add(1)
            module.CodeModule.AddFromString(
                """
                sub xl_paste()
                    ActiveSheet.Paste
                End sub
                """
            )
            app.api.Application.Run("xl_paste")
            win_wb.VBProject.VBComponents.Remove(module)
            try:
                wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AP Aging\\" +
                        str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                app.quit()
                print(facname + ' AP aging saved to shared drive')
            except:
                try:
                    os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                except:
                    pass
                try:
                    wb.save(userpath + '\\Desktop\\temp reporting\\' +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                    app.quit()
                    print(facname + ' AP aging saved to desktop 2')
                except:
                    print('Error saving AP aging to desktop')
                time.sleep(1)
        except:
            print('Issue downloading AP Aging: ' + facname)

    def ar_aging(self, facname, bu):
        """Download AR aging reports (saves Excel file)"""
        try:
            iter = True
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            title = self.driver.find_element(By.ID, "pccFacLink")
            time.sleep(1)
            if title.text != "Enterprise Management Console":
                iter = False
                self.driver.get("https://www30.pointclickcare.com/home/home.jsp")
                self.driver.find_element(By.ID, "pccFacLink").click()
                time.sleep(1)
                self.driver.find_element(By.CSS_SELECTOR, "#facTabs .pccButton").click()  # go to management console
                time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/emc/admin/reports/rp_araging_us.jsp")  # go to reports
            self.driver.find_element(By.LINK_TEXT, "select").click()  # click facilities
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.CSS_SELECTOR, "#footer > input:nth-child(2)").click()  # clear all
            self.driver.find_element(By.ID, "ESOLfacid_" + str(bu)).click()  # select building
            self.driver.find_element(By.CSS_SELECTOR, ".pccButton:nth-child(3)").click()  # save and exit
            self.driver.switch_to.window(window_before)  # go back to original tab.  Facility is selected
            if not iter:
                dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")  # select the reporting date
                dropdown.find_element(By.XPATH, "//option[. = \'" + prev_month_word + "\']").click()
                self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
                dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
                dropdown.select_by_value(str(report_year))
                self.driver.find_element(By.NAME, "ESOLyearSelect").click()
                self.driver.find_element(By.ID, "resdetailClient").click()
                self.driver.find_element(By.ID, "ESOLexport").click()
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(13)
            self.close_all_windows(window_before)
            try:
                convert_to_xlsx()  # change format from xls to xlsx
                try:
                    renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + " AR Aging",
                                         "P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AR Aging\\")
                except:
                    print('Issue moving and renaming the file')
            except:
                print('Issue converting excel file')
        except:
            print('Issue downloading AR Aging: ' + facname)

    def ar_rollforward(self, facname):
        """Download AR rollforward report (paste to Excel)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_arreconciliation_us.jsp")
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")
            dropdown.find_element(By.XPATH, "//option[. = \'" + prev_month_abbr + "\']").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.NAME, "ESOLyearSelect").click()
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            pyperclip.copy('')
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'a')  # highlight the entire page
            self.driver.find_element(By.CLASS_NAME, "admin").send_keys(Keys.CONTROL, 'a')
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # COPY ALL CONTENT
            self.driver.find_element(By.CLASS_NAME, "admin").send_keys(Keys.CONTROL, 'c')
            time.sleep(5)
            self.close_all_windows(window_before)
            wb = xw.Book()  # new workbook
            app = xw.apps.active
            win_wb = wb.api
            module = win_wb.VBProject.VBComponents.Add(1)
            module.CodeModule.AddFromString(
                """
                sub xl_paste()
                    ActiveSheet.Paste
                End sub
                """
            )
            app.api.Application.Run("xl_paste")
            win_wb.VBProject.VBComponents.Remove(module)
            try:
                wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AR Rollforward\\" +
                        str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                app.quit()
                print(facname + ' AR Rollforward saved to shared drive')
            except:
                try:
                    os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                except:
                    pass
                try:
                    wb.save(userpath + '\\Desktop\\temp reporting\\' +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                    app.quit()
                    print(facname + ' AR Rollforward saved to desktop 2')
                except:
                    print('Error saving AR Rollforward to desktop')
                time.sleep(2)
        except:
            print('Issue downloading AR Rollforward: ' + facname)

    def cash_receipts(self, facname):
        """Download cash receipts report (PDF)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_cashreceiptsjournal_us.jsp")
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLdateselect_active").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")
            dropdown.find_element(By.CSS_SELECTOR, "#pickdate > select:nth-child(2) > option:nth-child(" + str(
                prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            self.driver.find_element(By.NAME, "ESOLyearSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            time.sleep(5)
            html = self.driver.page_source
            while True:
                if "--PROCESSING--" not in html:
                    break
                else:
                    time.sleep(2)
            while True:
                try:
                    self.driver.execute_script('window.print();')  # print to PDF
                    break
                except:
                    time.sleep(2)
            time.sleep(3)
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Cash Receipts',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Cash Receipts\\')
        except:
            print('Issue downloading Cash Receipts: ' + facname)

    def census(self, facname):
        """Download census report (PDF)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get(
                "https://www30.pointclickcare.com/admin/reports/rp_detailedcensusWMY.jsp?ESOLfromER=Y&reportModule=P")
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "#summBySections label").click()
            self.driver.find_element(By.NAME, "ESOLmonth").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonth")
            dropdown.find_element(By.CSS_SELECTOR, "#periodspanid > select:nth-child(1) > option:nth-child(" + str(
                prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonth").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyear"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "ESOLsummByClient").click()
            self.driver.find_element(By.ID, "ESOLtotalByPayer").click()
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            time.sleep(5)
            html = self.driver.page_source
            while True:
                if "--PROCESSING--" not in html:
                    break
                else:
                    time.sleep(5)
            while True:
                try:
                    self.driver.execute_script('window.print();')  # print to PDF
                    break
                except:
                    time.sleep(5)
            time.sleep(5)
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Census',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Census\\')
        except:
            print('Issue downloading Census: ' + facname)

    def journal_entries(self, facname):
        """Download journal entries report (PDF)"""
        try:
            time.sleep(1)
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get(
                "https://www30.pointclickcare.com/admin/reports/rp_journalentries_fac_us.xhtml?action=setupReport")
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")
            dropdown.find_element(By.CSS_SELECTOR,
                                  "#postingPeriod > td:nth-child(3) > select:nth-child(2) > option:nth-child(" + str(
                                      prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(20)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Journal Entries',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Journal Entries\\')
        except:
            print('Issue downloading Journal Entries: ' + facname)

    def revenuerec(self, facname):
        """Download revenue reconciliation report (PDF)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_revenuerec_us.jsp")
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")
            dropdown.find_element(By.CSS_SELECTOR, "#postingControls > select:nth-child(1) > option:nth-child(" + str(
                prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            time.sleep(5)
            html = self.driver.page_source
            while True:
                if "--PROCESSING--" not in html:
                    break
                else:
                    time.sleep(5)
            while True:
                try:
                    self.driver.execute_script('window.print();')  # print to PDF
                    break
                except:
                    time.sleep(5)
            time.sleep(5)
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Revenue Reconciliation',
                'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Revenue Reconciliation\\')
        except:
            print('Issue downloading Revenue Reconciliation: ' + facname)


######################
#### GUI SECTION #####
######################


class MainWindow(QMainWindow):
    check_box = None
    tray_icon = None

    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(1000, 500))  # Set sizes
        self.setWindowTitle("PCC Reporting Program")  # Set a title
        self.central_widget = QWidget(self)  # Create a central widget
        self.setCentralWidget(self.central_widget)  # Set the central widget

        self.grid_layout = QGridLayout(self)  # Create a QGridLayout
        self.central_widget.setLayout(self.grid_layout)  # Set the layout into the central widget
        self.grid_layout.addWidget(QLabel("Welcome", self), 0, 0)

        self.report_button = QPushButton('Month End Reports', self)
        self.grid_layout.addWidget(self.report_button, 1, 0)
        self.report_button.clicked.connect(self.open_reports)

        self.status_box = QTextEdit()
        self.grid_layout.addWidget(self.status_box, 7, 0)
        self.setLayout(self.grid_layout)

    def open_reports(self):
        self.child_win = RunReportsWin()
        self.child_win.show()

    def update_textbox(self, message):
        self.status_box.append(f"{datetime.datetime.now().strftime('%I:%M:%S')}>>{message}")
        self.status_box.repaint()
        
    def closeEvent(self, event):
        """
        By overriding closeEvent, we can ignore the event and instead
        hide the window, effectively performing a "close-to-system-tray"
        action. To exit, the right-click->Exit option from the system
        tray must be used.
        """
        event.ignore()
        self.hide()
        self.notify("App minimized to system tray.")


class RunReportsWin(QWidget):
    def __init__(self):
        super(RunReportsWin, self).__init__()
        self.title = 'Select your buildings'
        self.left = 1200
        self.top = 200
        self.setWindowTitle(self.title)

        self.mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(self.mainframe)  # add the layout to the window

        self.cbframe = QFrame(self)  # frame that holds the check boxes
        self.cbframe.setFrameShape(QFrame.StyledPanel)  # add some style to the frame
        self.cbframe.setLineWidth(1)
        self.layout = QGridLayout(self.cbframe)  # create and add a layout for the frame
        self.mainframe.addWidget(self.cbframe)  # add the layout to the frame

        x, y = 1, 1  # add checkboxes to the layout of cbframe
        for item in facilities:  #
            cb = QCheckBox(str(item))  #
            cb.setChecked(False)  # set all checkboxes to unchecked
            self.layout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 10:  #
                x += 1  #
                y = 1  #

        self.rptframe = QFrame(self)  # create frame to select reports
        self.rptlayout = QGridLayout(self.rptframe)  # create and add grid layout to the frame
        self.mainframe.addWidget(self.rptframe)  # add frame to self.mainframe

        x, y = 1, 1
        for report in reports_list:  # create reports checkboxes
            cb = QCheckBox(report)  #
            cb.setChecked(False)  #
            self.rptlayout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 3:  #
                x += 1  #
                y = 1  #

        self.dateframe = QFrame(self)
        self.datelayout = QFormLayout(self.dateframe)
        self.mainframe.addWidget(self.dateframe)

        self.monthtextbox = QLineEdit(self)
        self.monthtextbox.setText(prev_month_num_str)
        self.monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', self.monthtextbox)
        self.yeartextbox = QLineEdit(self)
        self.yeartextbox.setText(str(report_year))
        self.yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', self.yeartextbox)

        self.btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        self.btnlayout = QGridLayout(self.btnframe)  # create and add a layout for the frame
        self.mainframe.addWidget(self.btnframe)  # add the frame to the main frame

        self.saverunbtn = QPushButton('Save and Run', self)
        self.btnlayout.addWidget(self.saverunbtn, 1, 1)
        self.saverunbtn.clicked.connect(self.checkCheckboxes)

        self.selectallbtn = QPushButton('Check All', self)
        self.btnlayout.addWidget(self.selectallbtn, 1, 2)
        self.selectallbtn.clicked.connect(self.selectCheckboxes)

        self.unselectallbtn = QPushButton('Uncheck All', self)
        self.btnlayout.addWidget(self.unselectallbtn, 1, 3)
        self.unselectallbtn.clicked.connect(self.unselectCheckboxes)

        self.checkrptsbtn = QPushButton('Run All', self)
        self.btnlayout.addWidget(self.checkrptsbtn, 1, 4)
        self.checkrptsbtn.clicked.connect(self.reportCounter)

    def checkCheckboxes(self):
        fac_checked_list = []
        rpt_checked_list = []

        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            if chbox.isChecked():
                fac_checked_list.append(chbox.text())

        for i in range(self.rptlayout.count()):
            chbox = self.rptlayout.itemAt(i).widget()
            if chbox.isChecked():
                rpt_checked_list.append(chbox.text())

        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        download_reports(fac_checked_list, rpt_checked_list)

    def selectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(True)

    def unselectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(False)

    def reportCounter(self):
        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        check_reports()
        # t1 = Thread(target=check_reports)
        # t1.start()
        # while True:
        #     if not q.empty():
        #         data = q.get()
        #         print(data)
        #         if event.is_set():
        #             break
        #     else:
        #         time.sleep(0.01)
        # t1.join()
        # print("Finished")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
