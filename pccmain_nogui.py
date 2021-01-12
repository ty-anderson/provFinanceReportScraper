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
    QPushButton, QVBoxLayout, QFrame, QFormLayout, QLineEdit
from PyQt5.QtCore import QSize
import sys
import pyperclip

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
global check_status

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
accountantlist = facility_df['Accountant'].drop_duplicates().to_list()  # make list of all accountants
reports_list = ['AP Aging',
                'AR Aging',
                'AR Rollforward',
                'Cash Receipts Journal',
                'Detailed Census',
                'Journal Entries',
                'Revenue Reconciliation']


def to_text(message):
    """Write to text file to notify user"""
    s = str(datetime.datetime.now().strftime("%H:%M:%S")) + ">>  " + str(message) + "\n"
    with open(userpath + '\\Desktop\\PyReport.txt', 'a') as file:
        file.write(s)
        file.close()


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
    to_text('Reporting date is ' + prev_month_abbr + ' ' + str(report_year))


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
            to_text('Moved to: ' + destfile)  # END BACKUP LOCATION
    except:
        to_text("Issue renaming/moving to " + str(dirpath))
        to_text(newfilename + " is in Downloads folder")


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
    if report == "Cash Receipts Journal":
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
    elif report == "Detailed Census":
        report_name = "Census.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census'
    elif report == "Journal Entries":
        report_name = "Journal Entries.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries'
    elif report == "Revenue Reconciliation":
        report_name = "Revenue Reconciliation.pdf"
        report_path = r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation'
    else:
        report_name = "Issue identifying report"
        report_path = "Issue identifying report"
    file_name = report_path + '\\' + str(report_year) + ' ' + str(prev_month_num_str) + ' ' + facility + ' ' + report_name
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
            to_text('Updating chromedriver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + max(file_list) + '.exe')
            to_text('chromedriver updated to version ' + max(file_list))
        except:
            to_text("Couldn't update chromedriver automatically")
        return max(file_list)
    else:
        to_text('Could not find P:\\PACS\\Finance\\Automation\\Chromedrivers\\')


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


def downloadKindredReport():
    """Download census reports for Kindred buildings"""
    to_text("Downloading weekly census report for Kindred buildings")
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()
    today = datetime.date.today()
    if today.weekday() == 6:
        sunday = datetime.date.today()
    else:
        sunday = (today + datetime.timedelta(days=(-today.weekday() - 1), weeks=0))  # gets previous monday
    sundaystr = str(sunday.month) + "-" + str(sunday.day) + "-" + str(sunday.year)
    PCC.kindredReport()
    wb = xw.Book()  # new workbook
    time.sleep(3)
    wb.activate(steal_focus=True)  # focus the new instance
    pyautogui.hotkey('ctrl', 'v')  # paste
    time.sleep(2)  # wait to load
    xw.Range('F:Z').delete()  # clear the unused columns
    try:
        wb.save(userpath + '\\Documents\\Kindred Reporting\\Week Ending ' + sundaystr + '.xlsx')
        wb.close()
        time.sleep(2)
        os.startfile(userpath + '\\Documents\\Kindred Reporting\\Week Ending ' + sundaystr + '.xlsx')
    except FileNotFoundError:
        to_text('Workbook was not saved.')
    to_text('Process has finished')


def downloadIncomeStmtM2M(facilitylist):
    """Download income statements for prelims"""
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    deleteDownloads()
    for facname in facilities:
        bu = str(facilities[facname][1])
        if len(bu) < 2:
            bu = str(0) + bu
        if facname in facilitylist:
            to_text("Downloading income statement: " + facname)
            if PCC.buildingSelect(bu):  # get next building in the list on chrome
                time.sleep(1)
                PCC.IS_M2M(str(report_year), facname)  # run reports
    PCC.teardown_method()
    to_text("Income statements downloaded")


def downloadIntercoReports():
    """Download intercompany reports needed to reconcile"""
    deleteDownloads()
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    to_text('Downloading intercompany reports')
    PCC.intercompany_reports()
    PCC.teardown_method()
    to_text('Complete')


def downloadTrustReports(facilitylist):
    """Download reports to reconcile trust per building"""
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    for facname in facilities:
        if facname in facilitylist:
            bu = str(facilities[facname][1])
            if len(bu) < 2:
                bu = str(0) + bu
            if PCC.buildingSelect(bu):
                time.sleep(1)
                PCC.trust_reports()


def downloadAuditReports(facilitylist):
    """Download reports to reconcile trust per building"""
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    for facname in facilities:
        if facname in facilitylist:
            bu = str(facilities[facname][1])
            if len(bu) < 2:
                bu = str(0) + bu
            if PCC.buildingSelect(bu):
                time.sleep(1)
                PCC.ar_credit_balances(facname)


# run the reports
def download_reports(facilitylist=facilityindex, reportlist=reports_list):
    """Download month end close reports"""
    global check_status
    global PCC
    deleteDownloads()
    counter = 0
    if not facilitylist:
        facilitylist = facilityindex
    if reportlist:
        check_status = True
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
                        if check_status:
                            if report == 'AP Aging':
                                PCC.ap_aging(facname)
                            if report == 'AR Aging':  # USES MGMT CONSOLE
                                bu = facilities[facname][1]  # TO SELECT BUILDING IN AR REPORT
                                PCC.ar_aging(facname, bu)
                                PCC.buildingSelect(str(bu))
                            if report == 'AR Rollforward':
                                PCC.ar_rollforward(facname)
                            if report == 'Cash Receipts Journal':
                                PCC.cash_receipts(facname)
                            if report == 'Detailed Census':
                                PCC.census(facname)
                            if report == 'Journal Entries':
                                PCC.journal_entries(facname)
                            if report == 'Revenue Reconciliation':
                                PCC.revenuerec(facname)
                            counter += 1
                            check_if_downloaded(facname, report)
                        else:
                            to_text('There is an issue with the chromedriver')
        to_text('Reports downloaded')
        # PCC.teardown_method()
        # del PCC
    else:
        to_text('No reports selected.')


class LoginPCC:
    def __init__(self):
        """Create instance, login to PCC"""
        global check_status
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
            try:
                chromedriver_autoinstaller.install()
                self.driver = webdriver.Chrome(options=chrome_options)
            except:
                try:
                    latestdriver = find_current_driver()
                    self.driver = webdriver.Chrome(
                        os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                        options=chrome_options)
                    to_text('Local chromedriver successfully initiated')
                except:
                    latestdriver = find_updated_driver()
                    self.driver = webdriver.Chrome(
                        os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                        options=chrome_options)
                    to_text('Chromedrive successfully initiated')
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
            to_text('There was an issue initiating chromedriver')
            check_status = False

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
            current_fac = self.driver.find_element(By.NAME, "current_fac_id").get_attribute("value")
            if str(current_fac) != bu:
                try:
                    self.driver.find_element(By.ID, "pccFacLink").click()
                    time.sleep(1)
                    building_list = self.driver.find_element(By.ID, "optionList")
                    building_list.find_element(By.PARTIAL_LINK_TEXT, bu).click()
                    return True
                except:
                    to_text("Could not locate " + bu + " in PCC")
                    return False
            else:
                return True
        except:
            to_text("Could not find the building dropdown menu")
            return False

    def IS_M2M(self, year, facname):
        """Download income statement M-to-M report (download Excel file)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get('https://www30.pointclickcare.com/glap/reports/rp_customglreports.jsp?ESOLrepId=555')
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLyear").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLyear")
            dropdown.find_element(By.XPATH, "//option[. = " + str(year) + "]").click()
            self.driver.find_element(By.NAME, "ESOLdispAcctNo").click()
            self.driver.find_element(By.NAME, "ESOLshowDecimals").click()
            self.driver.find_element(By.NAME, "ESOLExportToSpreadsheet").click()
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(10)
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.CSS_SELECTOR, ".pccButton:nth-child(1)").click()
            time.sleep(10)
            # self.driver.close()
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(prev_month_num) + " " + str(report_year) + " " + facname + ' Income Statement M-to-M',
                userpath + "\\Documents\\AutoFillFinancials\\")  # rename and move file
        except:
            to_text('There was an issue downloading')

    def ap_aging(self, facname):
        """Download AP aging report (paste to Excel)"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/glap/reports/rp_aptrialbalance.xhtml")
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "tr:nth-child(3) label:nth-child(3) > input").click()
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
            self.driver.find_element(By.CLASS_NAME, "admin").send_keys(Keys.CONTROL, 'a')
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # COPY ALL CONTENT
            self.driver.find_element(By.CLASS_NAME, "admin").send_keys(Keys.CONTROL, 'c')
            time.sleep(2)
            self.close_all_windows(window_before)
            wb = xw.Book()  # new workbook
            app = xw.apps.active
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
            try:
                wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AP Aging\\" +
                        str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                app.quit()
                to_text(facname + ' AP aging saved to shared drive')
            except:
                try:
                    os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                except:
                    pass
                try:
                    wb.save(userpath + '\\Desktop\\temp reporting\\' +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                    app.quit()
                    to_text(facname + ' AP aging saved to desktop 2')
                except:
                    to_text('Error saving AP aging to desktop')
                time.sleep(2)
        except:
            to_text('Issue downloading AP Aging: ' + facname)

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
                    to_text('Issue moving and renaming the file')
            except:
                to_text('Issue converting excel file')
        except:
            to_text('Issue downloading AR Aging: ' + facname)

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
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
            try:
                wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AR Rollforward\\" +
                        str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                app.quit()
                to_text(facname + ' AR Rollforward saved to shared drive')
            except:
                try:
                    os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                except:
                    pass
                try:
                    wb.save(userpath + '\\Desktop\\temp reporting\\' +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                    app.quit()
                    to_text(facname + ' AR Rollforward saved to desktop 2')
                except:
                    to_text('Error saving AR Rollforward to desktop')
                time.sleep(2)
        except:
            to_text('Issue downloading AR Rollforward: ' + facname)

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
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.execute_script('window.print();')  # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Cash Receipts',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Cash Receipts\\')
        except:
            to_text('Issue downloading Cash Receipts: ' + facname)

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
            self.driver.execute_script('window.print();')  # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Census',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Census\\')
        except:
            to_text('Issue downloading Census: ' + facname)

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
            time.sleep(10)  # wait
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Journal Entries',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Journal Entries\\')
        except:
            to_text('Issue downloading Journal Entries: ' + facname)

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
            self.driver.execute_script('window.print();')  # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Revenue Reconciliation',
                'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Revenue Reconciliation\\')
        except:
            to_text('Issue downloading Revenue Reconciliation: ' + facname)

    def close_ap_periods(self):
        """Close AP periods (not working yet)"""
        try:
            print("Closing the month of " + str(prev_month_abbr))
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get(
                'https://www30.pointclickcare.com/glap/setup/fiscalyearslist.jsp?ESOLrefer=https://www30.pointclickcare.com/glap/setup/glapsetup.jsp')
            self.driver.find_element(By.CSS_SELECTOR,
                                     "#expandoTabDivAP > div > table > tbody > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)").click()
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.XPATH, "//*[@id=\"fiscal" + str(prev_month_num) + "c\"]").click()
            self.driver.find_element(By.XPATH, "//*[@id=\"fiscal" + str(prev_month_num + 2) + "o\"]").click()
            self.driver.find_element(By.CSS_SELECTOR, "#msg > input:nth-child(1)").click()  # click button weekly
            time.sleep(2)
            self.driver.switch_to.window(window_before)  # go back to original tab
        except:
            to_text('There was an issue downloading')

    def close_gl_periods(self):
        """Close GL periods (not working yet)"""
        try:
            # print("Closing the month of " + str(prev_month_abbr))
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get(
                'https://www30.pointclickcare.com/glap/setup/fiscalyearslist.jsp?ESOLrefer=https://www30.pointclickcare.com/glap/setup/glapsetup.jsp')
            self.driver.find_element(By.CSS_SELECTOR,
                                     "#expandoTabDivGL > div > table > tbody > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)").click()
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.XPATH, "//*[@id=\"fiscal" + str(prev_month_num) + "c\"]").click()
            self.driver.find_element(By.CSS_SELECTOR, "#msg > input:nth-child(1)").click()  # click button weekly
            time.sleep(2)
            self.driver.switch_to.window(window_before)  # go back to original tab
        except:
            to_text('There was an issue downloading')

    def kindredReport(self):
        """Download Kindred census report"""
        try:
            tdy = datetime.date.today()
            if tdy.weekday() == 6:
                sunday = datetime.date.today()
            else:
                sunday = (tdy + datetime.timedelta(days=(-tdy.weekday() - 1), weeks=0))  # gets previous monday
            sundaystr = str(sunday.month) + "/" + str(sunday.day) + "/" + str(sunday.year)
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.find_element(By.ID, "pccFacLink").click()
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "#facTabs .pccButton").click()
            time.sleep(1)
            self.driver.get(
                "https://www30.pointclickcare.com/admin/reports/rp_detailedcensusWMY.jsp?allowEMCModeCheck=true")
            self.driver.find_element(By.CSS_SELECTOR,
                                     "tr:nth-child(3) label:nth-child(1)").click()  # click button weekly
            self.driver.find_element(By.CSS_SELECTOR, ".groupBy label:nth-child(2) > input").click()  # click facilities
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.ID, "ESOLfacid_58").click()  # Pac Coast building
            self.driver.find_element(By.ID, "ESOLfacid_59").click()  # Medical Hill building
            self.driver.find_element(By.ID, "ESOLfacid_60").click()  # Santa Cruz building
            self.driver.find_element(By.ID, "ESOLfacid_57").click()  # Santa Cruz building
            self.driver.find_element(By.CSS_SELECTOR, ".pccButton:nth-child(3)").click()
            self.driver.switch_to.window(window_before)  # go back to original tab
            self.driver.find_element(By.ID, "ESOLperiodend_dummy").click()
            self.driver.find_element(By.ID, "ESOLperiodend_dummy").send_keys(6 * Keys.BACKSPACE)
            self.driver.find_element(By.ID, "ESOLperiodend_dummy").send_keys(4 * Keys.DELETE)
            # self.driver.find_element(By.ID, "ESOLperiodend_dummy").clear()
            self.driver.find_element(By.ID, "ESOLperiodend_dummy").send_keys(sundaystr)
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(4)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'a')  # highlight the entire page
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # copy the entire page
            time.sleep(1)
            # self.driver.close()
            # self.driver.switch_to.window(window_before)  # go back to original tab
        except:
            to_text('There was an issue downloading')

    def intercompany_reports(self):
        """Downlaod intercompany reports to reconcile accounts"""
        window_before = self.driver.window_handles[0]  # make window tab object
        time.sleep(1)
        title = self.driver.find_element(By.ID, "pccFacLink")
        time.sleep(1)
        if title.text != "Enterprise Management Console":
            self.driver.get("https://www30.pointclickcare.com/home/home.jsp")
            self.driver.find_element(By.ID, "pccFacLink").click()
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "#facTabs .pccButton").click()  # go to management console
            time.sleep(1)
        self.driver.get("https://www30.pointclickcare.com/glap/reports/rp_gltransactions.xhtml")  # GL transactions
        time.sleep(4)
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLperstart"))  # month selector
        dropdown.select_by_value(str(prev_month_num))
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLyrstart"))
        dropdown.select_by_value(str(report_year))
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLperend"))
        dropdown.select_by_value(str(prev_month_num))
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLyrend"))
        dropdown.select_by_value(str(report_year))
        time.sleep(1)
        self.driver.find_element(By.CSS_SELECTOR,
                                 "body > table:nth-child(15) > tbody > tr:nth-child(19) > td:nth-child(3) > input[type=radio]:nth-child(10)").click()
        self.driver.find_element(By.CSS_SELECTOR,
                                 "body > table:nth-child(15) > tbody > tr:nth-child(19) > td:nth-child(3) > input[type=text]:nth-child(11)").send_keys(
            "1340.000")
        time.sleep(1)
        self.driver.find_element(By.CSS_SELECTOR,
                                 "body > table:nth-child(15) > tbody > tr:nth-child(19) > td:nth-child(3) > a > img").click()
        time.sleep(1)
        self.driver.find_element(By.CSS_SELECTOR,
                                 "body > table:nth-child(15) > tbody > tr:nth-child(19) > td:nth-child(3) > a > img").click()
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLreportOutputType"))
        time.sleep(2)
        dropdown.select_by_value('csv')
        time.sleep(2)
        self.driver.find_element(By.ID, "runButton").click()
        window_after = self.driver.window_handles[1]
        self.driver.switch_to.window(window_after)
        while True:
            try:
                self.driver.find_element(By.CSS_SELECTOR, "#ajaxComplete > td > input").click()
                self.driver.switch_to.window(window_before)
                break
            except:
                time.sleep(5)
        time.sleep(10)
        self.close_all_windows(window_before)  # end of GL transactions
        renameDownloadedFile("PCC Interco.csv")
        # download balance sheet
        self.driver.get("https://www30.pointclickcare.com/glap/reports/rp_customglreports.jsp?ESOLrepId=5")
        self.driver.find_element(By.CSS_SELECTOR, "#dateRange > table > tbody > tr > td:nth-child(" + str(
            prev_month_num) + ") > input[type=checkbox]:nth-child(3)").click()
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLyear"))
        dropdown.select_by_value(str(report_year))
        self.driver.find_element(By.NAME, "ESOLcomparefacs").click()
        self.driver.find_element(By.NAME, "ESOLdispAcctNo").click()
        self.driver.find_element(By.NAME, "ESOLshowDecimals").click()
        self.driver.find_element(By.NAME, "ESOLExportToSpreadsheet").click()
        self.driver.find_element(By.ID, "runButton").click()
        window_after = self.driver.window_handles[1]
        self.driver.switch_to.window(window_after)
        while True:
            try:
                self.driver.find_element(By.CSS_SELECTOR,
                                         "#ExportDiv > form > table > tbody > tr:nth-child(2) > td > input:nth-child(1)").click()
                self.driver.switch_to.window(window_before)
                break
            except:
                time.sleep(5)
        time.sleep(10)
        self.close_all_windows(window_before)

    def trust_reports(self):
        """Open resident trust reports"""
        window_before = self.driver.window_handles[0]  # make window tab object
        time.sleep(1)
        self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_ta_audit.jsp")  # audit report
        self.driver.find_element(By.NAME, "ESOLstartdate").click()
        self.driver.find_element(By.NAME, "ESOLstartdate").send_keys(6 * Keys.BACKSPACE)
        self.driver.find_element(By.NAME, "ESOLstartdate").send_keys(6 * Keys.DELETE)
        self.driver.find_element(By.NAME, "ESOLstartdate").send_keys(
            "{}/01/{}".format(prev_month_num_str, str(report_year)))
        self.driver.find_element(By.NAME, "ESOLenddate").click()
        self.driver.find_element(By.NAME, "ESOLenddate").send_keys(6 * Keys.BACKSPACE)
        self.driver.find_element(By.NAME, "ESOLenddate").send_keys(6 * Keys.DELETE)
        self.driver.find_element(By.NAME, "ESOLenddate").send_keys(
            "{}/31/{}".format(prev_month_num_str, str(report_year)))
        self.driver.find_element(By.ID, "runButton").click()
        self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_ta_acct_bal.jsp")  # account balances
        self.driver.find_element(By.NAME, "ESOLfromdate").click()
        self.driver.find_element(By.NAME, "ESOLfromdate").send_keys(6 * Keys.BACKSPACE)
        self.driver.find_element(By.NAME, "ESOLfromdate").send_keys(6 * Keys.DELETE)
        self.driver.find_element(By.NAME, "ESOLfromdate").send_keys(
            "{}/31/{}".format(prev_month_num_str, str(report_year)))
        self.driver.find_element(By.ID, "runButton").send_keys(Keys.TAB)
        self.driver.find_element(By.ID, "runButton").click()
        self.driver.get(
            "https://www30.pointclickcare.com/admin/reports/rp_detailedcensusWMY.jsp?ESOLfromER=Y&reportModule=P")  # census
        dropdown = self.driver.find_element(By.NAME, "ESOLmonth")
        dropdown.find_element(By.CSS_SELECTOR, "#periodspanid > select:nth-child(1) > option:nth-child(" + str(
            prev_month_num) + ")").click()
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLyear"))
        dropdown.select_by_value(str(report_year))
        self.driver.find_element(By.ID, "runButton").click()
        self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_cashreceiptsjournal_us.jsp")  # Cash receipts
        self.driver.find_element(By.NAME, "ESOLdateselect_active").click()
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLmonthSelect"))
        dropdown.select_by_value(str(prev_month_num))
        dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
        dropdown.select_by_value(str(report_year))
        self.driver.find_element(By.ID, "runButton").click()

    def ar_credit_balances(self, facname):
        """Pull reports for CapFund audit"""
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www30.pointclickcare.com/admin/reports/rp_araging_us.jsp")
            time.sleep(1)
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLmonthSelect"))
            dropdown.select_by_value(str(prev_month_num))
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.execute_script('window.print();')  # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Aging Credit Balances',
                'P:\\PACS\\Finance\\Audit\\CapFund Bank Audit Jan 2021\\Accounts Receivable\\5 - Aged AR Credit Balances\\')
        except:
            to_text('Issue downloading: ' + facname)


# GUI SECTION *************************************************************************************************


class MainWindow(QMainWindow):
    check_box = None
    tray_icon = None

    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(380, 200))  # Set sizes
        self.setWindowTitle("PCC Reporting Program")  # Set a title
        central_widget = QWidget(self)  # Create a central widget
        self.setCentralWidget(central_widget)  # Set the central widget

        grid_layout = QGridLayout(self)  # Create a QGridLayout
        central_widget.setLayout(grid_layout)  # Set the layout into the central widget
        grid_layout.addWidget(QLabel("Welcome", self), 0, 0)

        self.report_button = QPushButton('Month End Reports', self)
        grid_layout.addWidget(self.report_button, 1, 0)
        self.report_button.clicked.connect(self.open_reports)

        self.incomestmt_button = QPushButton('Income Statements', self)
        grid_layout.addWidget(self.incomestmt_button, 2, 0)
        self.incomestmt_button.clicked.connect(self.open_incomestmt)

        self.kindred_button = QPushButton('Kindred Report', self)
        grid_layout.addWidget(self.kindred_button, 3, 0)
        self.kindred_button.clicked.connect(downloadKindredReport)

        self.interco_button = QPushButton('Intercompany Reports', self)
        grid_layout.addWidget(self.interco_button, 4, 0)
        self.interco_button.clicked.connect(self.open_intercowin)

        self.trust_button = QPushButton('Resident Trust', self)
        grid_layout.addWidget(self.trust_button, 5, 0)
        self.trust_button.clicked.connect(self.open_trustwin)

        self.trust_button = QPushButton('Audit Reports', self)
        grid_layout.addWidget(self.trust_button, 6, 0)
        self.trust_button.clicked.connect(self.open_auditwin)

        # status box
        # self.status_box = QTextBrowser(self)
        # grid_layout.addWidget(self.status_box, 6, 0)
        # self.text_out("Hello")


    def open_reports(self):
        self.child_win = RunReportsWin()
        self.child_win.show()

    def open_incomestmt(self):
        self.child_win2 = RunIncomeStmtWin()
        self.child_win2.show()

    def open_intercowin(self):
        self.child_win2 = RunIntercoWin()
        self.child_win2.show()

    def open_trustwin(self):
        self.child_win2 = RunTrustWin()
        self.child_win2.show()

    def open_auditwin(self):
        self.child_win2 = RunAuditWin()
        self.child_win2.show()

    def kill_program(self):
        exit()


class RunReportsWin(QWidget):
    def __init__(self):
        super(RunReportsWin, self).__init__()
        self.title = 'Select your buildings'
        self.left = 1200
        self.top = 200
        # self.width = 520
        # self.height = 400
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)

        mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(mainframe)  # add the layout to the window

        self.cbframe = QFrame(self)  # frame that holds the check boxes
        self.cbframe.setFrameShape(QFrame.StyledPanel)  # add some style to the frame
        self.cbframe.setLineWidth(0.6)
        self.layout = QGridLayout(self.cbframe)  # create and add a layout for the frame
        mainframe.addWidget(self.cbframe)  # add the layout to the frame

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
        mainframe.addWidget(self.rptframe)  # add frame to mainframe

        x, y = 1, 1
        for report in reports_list:  # create reports checkboxes
            cb = QCheckBox(report)  #
            cb.setChecked(False)  #
            self.rptlayout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 3:  #
                x += 1  #
                y = 1  #

        dateframe = QFrame(self)
        self.datelayout = QFormLayout(dateframe)
        mainframe.addWidget(dateframe)

        monthtextbox = QLineEdit(self)
        monthtextbox.setText(prev_month_num_str)
        monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', monthtextbox)
        yeartextbox = QLineEdit(self)
        yeartextbox.setText(str(report_year))
        yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', yeartextbox)

        btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        btnlayout = QGridLayout(btnframe)  # create and add a layout for the frame
        mainframe.addWidget(btnframe)  # add the frame to the main frame

        saverunbtn = QPushButton('Save and Run', self)
        btnlayout.addWidget(saverunbtn, 1, 1)
        saverunbtn.clicked.connect(self.checkCheckboxes)
        selectallbtn = QPushButton('Check All', self)
        btnlayout.addWidget(selectallbtn, 1, 2)
        selectallbtn.clicked.connect(self.selectCheckboxes)
        unselectallbtn = QPushButton('Uncheck All', self)
        btnlayout.addWidget(unselectallbtn, 1, 3)
        unselectallbtn.clicked.connect(self.reportCounter)
        unselectallbtn = QPushButton('Count Reports', self)
        btnlayout.addWidget(unselectallbtn, 1, 4)
        unselectallbtn.clicked.connect(self.reportCounter)

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
        wb_ref = r"P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx"
        wb = pd.read_excel(wb_ref, sheet_name='Automation', usecols=['Common Name'])
        wb_list = wb['Common Name'].to_list()
        reports_path = [r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AP Aging',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries',
                        r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation']
        report_names = ['AP Aging.xlsx', 'AR Aging.xlsx', 'AR Rollforward.xlsx', 'Cash Receipts.pdf',
                        'Census.pdf', 'Journal Entries.pdf', 'Revenue Reconciliation.pdf']
        i = 0
        for path in reports_path:
            for building in wb_list:
                file_name = path + '\\' + str(report_year) + ' ' + str(prev_month_num_str) + ' ' + building + ' ' + report_names[i]
                if not os.path.exists(file_name):
                    print(file_name + ' missing.  Downloading now')
                    rpt = report_names[i].split('.')
                    rpt = [rpt[0]]
                    download_reports(building, rpt)
            i+=1


class RunIncomeStmtWin(QWidget):
    def __init__(self):
        super(RunIncomeStmtWin, self).__init__()
        self.title = 'Select your buildings'
        self.left = 1200
        self.top = 200
        # self.width = 520
        # self.height = 400
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)

        mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(mainframe)  # add the layout to the window

        self.cbframe = QFrame(self)  # frame that holds the check boxes
        self.cbframe.setFrameShape(QFrame.StyledPanel)  # add some style to the frame
        self.cbframe.setLineWidth(0.6)
        self.layout = QGridLayout(self.cbframe)  # create and add a layout for the frame
        mainframe.addWidget(self.cbframe)  # add the layout to the frame

        x, y = 1, 1  # add checkboxes to the layout of cbframe
        for item in facilities:  #
            cb = QCheckBox(str(item))  #
            cb.setChecked(False)  # set all checkboxes to unchecked
            self.layout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 10:  #
                x += 1  #
                y = 1  #

        dateframe = QFrame(self)
        self.datelayout = QFormLayout(dateframe)
        mainframe.addWidget(dateframe)

        monthtextbox = QLineEdit(self)
        monthtextbox.setText(prev_month_num_str)
        monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', monthtextbox)
        yeartextbox = QLineEdit(self)
        yeartextbox.setText(str(report_year))
        yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', yeartextbox)

        btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        btnlayout = QGridLayout(btnframe)  # create and add a layout for the frame
        mainframe.addWidget(btnframe)  # add the frame to the main frame

        saverunbtn = QPushButton('Save and Run', self)
        btnlayout.addWidget(saverunbtn, 1, 1)
        saverunbtn.clicked.connect(self.checkCheckboxes)
        selectallbtn = QPushButton('Check All', self)
        btnlayout.addWidget(selectallbtn, 1, 2)
        selectallbtn.clicked.connect(self.selectCheckboxes)
        unselectallbtn = QPushButton('Uncheck All', self)
        btnlayout.addWidget(unselectallbtn, 1, 3)
        unselectallbtn.clicked.connect(self.unselectCheckboxes)

    def checkCheckboxes(self):
        fac_checked_list = []
        rpt_checked_list = []

        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            if chbox.isChecked():
                fac_checked_list.append(chbox.text())

        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        downloadIncomeStmtM2M(fac_checked_list)

    def selectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(True)

    def unselectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(False)


class RunIntercoWin(QWidget):
    def __init__(self):
        super(RunIntercoWin, self).__init__()
        self.title = 'Select date'
        self.left = 1200
        self.top = 200
        # self.width = 520
        # self.height = 400
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)

        mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(mainframe)  # add the layout to the window

        dateframe = QFrame(self)
        self.datelayout = QFormLayout(dateframe)
        mainframe.addWidget(dateframe)

        monthtextbox = QLineEdit(self)
        monthtextbox.setText(prev_month_num_str)
        monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', monthtextbox)
        yeartextbox = QLineEdit(self)
        yeartextbox.setText(str(report_year))
        yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', yeartextbox)

        btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        btnlayout = QGridLayout(btnframe)  # create and add a layout for the frame
        mainframe.addWidget(btnframe)  # add the frame to the main frame

        saverunbtn = QPushButton('Save and Run', self)
        btnlayout.addWidget(saverunbtn, 1, 1)
        saverunbtn.clicked.connect(self.runInterco)

    def runInterco(self):
        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        downloadIntercoReports()


class RunTrustWin(QWidget):
    def __init__(self):
        super(RunTrustWin, self).__init__()
        self.title = 'Select your buildings'
        self.left = 1200
        self.top = 200
        # self.width = 520
        # self.height = 400
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)

        mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(mainframe)  # add the layout to the window

        self.cbframe = QFrame(self)  # frame that holds the check boxes
        self.cbframe.setFrameShape(QFrame.StyledPanel)  # add some style to the frame
        self.cbframe.setLineWidth(0.6)
        self.layout = QGridLayout(self.cbframe)  # create and add a layout for the frame
        mainframe.addWidget(self.cbframe)  # add the layout to the frame

        x, y = 1, 1  # add checkboxes to the layout of cbframe
        for item in facilities:  #
            cb = QCheckBox(str(item))  #
            cb.setChecked(False)  # set all checkboxes to unchecked
            self.layout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 10:  #
                x += 1  #
                y = 1  #

        dateframe = QFrame(self)
        self.datelayout = QFormLayout(dateframe)
        mainframe.addWidget(dateframe)

        monthtextbox = QLineEdit(self)
        monthtextbox.setText(prev_month_num_str)
        monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', monthtextbox)
        yeartextbox = QLineEdit(self)
        yeartextbox.setText(str(report_year))
        yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', yeartextbox)

        btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        btnlayout = QGridLayout(btnframe)  # create and add a layout for the frame
        mainframe.addWidget(btnframe)  # add the frame to the main frame

        saverunbtn = QPushButton('Save and Run', self)
        btnlayout.addWidget(saverunbtn, 1, 1)
        saverunbtn.clicked.connect(self.checkCheckboxes)
        selectallbtn = QPushButton('Check All', self)
        btnlayout.addWidget(selectallbtn, 1, 2)
        selectallbtn.clicked.connect(self.selectCheckboxes)
        unselectallbtn = QPushButton('Uncheck All', self)
        btnlayout.addWidget(unselectallbtn, 1, 3)
        unselectallbtn.clicked.connect(self.unselectCheckboxes)

    def checkCheckboxes(self):
        fac_checked_list = []
        rpt_checked_list = []

        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            if chbox.isChecked():
                fac_checked_list.append(chbox.text())

        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        downloadTrustReports(fac_checked_list)

    def selectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(True)

    def unselectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(False)


class RunAuditWin(QWidget):
    def __init__(self):
        super(RunAuditWin, self).__init__()
        self.title = 'Select your buildings'
        self.left = 1200
        self.top = 200
        # self.width = 520
        # self.height = 400
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)

        mainframe = QVBoxLayout()  # create a layout for the window
        self.setLayout(mainframe)  # add the layout to the window

        self.cbframe = QFrame(self)  # frame that holds the check boxes
        self.cbframe.setFrameShape(QFrame.StyledPanel)  # add some style to the frame
        self.cbframe.setLineWidth(0.6)
        self.layout = QGridLayout(self.cbframe)  # create and add a layout for the frame
        mainframe.addWidget(self.cbframe)  # add the layout to the frame

        x, y = 1, 1  # add checkboxes to the layout of cbframe
        for item in facilities:  #
            cb = QCheckBox(str(item))  #
            cb.setChecked(False)  # set all checkboxes to unchecked
            self.layout.addWidget(cb, y, x)  #
            y += 1  #
            if y >= 10:  #
                x += 1  #
                y = 1  #

        dateframe = QFrame(self)
        self.datelayout = QFormLayout(dateframe)
        mainframe.addWidget(dateframe)

        monthtextbox = QLineEdit(self)
        monthtextbox.setText(prev_month_num_str)
        monthtextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Month:', monthtextbox)
        yeartextbox = QLineEdit(self)
        yeartextbox.setText(str(report_year))
        yeartextbox.setFixedSize(100, 20)
        self.datelayout.addRow('Year:', yeartextbox)

        btnframe = QFrame(self)  # create a new frame for save and run, check all, uncheck all
        btnlayout = QGridLayout(btnframe)  # create and add a layout for the frame
        mainframe.addWidget(btnframe)  # add the frame to the main frame

        saverunbtn = QPushButton('Save and Run', self)
        btnlayout.addWidget(saverunbtn, 1, 1)
        saverunbtn.clicked.connect(self.checkCheckboxes)
        selectallbtn = QPushButton('Check All', self)
        btnlayout.addWidget(selectallbtn, 1, 2)
        selectallbtn.clicked.connect(self.selectCheckboxes)
        unselectallbtn = QPushButton('Uncheck All', self)
        btnlayout.addWidget(unselectallbtn, 1, 3)
        unselectallbtn.clicked.connect(self.unselectCheckboxes)

    def checkCheckboxes(self):
        fac_checked_list = []
        rpt_checked_list = []

        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            if chbox.isChecked():
                fac_checked_list.append(chbox.text())

        month = self.datelayout.itemAt(1).widget()
        year = self.datelayout.itemAt(3).widget()
        update_date(month.text(), year.text())
        self.close()
        downloadAuditReports(fac_checked_list)

    def selectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(True)

    def unselectCheckboxes(self):
        for i in range(self.layout.count()):
            chbox = self.layout.itemAt(i).widget()
            chbox.setChecked(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())
