import calendar
import shutil
import glob
import pandas as pd
from tkinter import *
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
import csv
import win32com
import threading
from infi.systray import SysTrayIcon


# clear the gen_py folder that is causing issues with the xlsx conversion with win32com
try:
    shutil.rmtree(win32com.__gen_path__[:-4])
except:
    pass

global newpathtext

# collect user info
username = os.environ['USERNAME']
userpath = os.environ['USERPROFILE']

# collect date info
today = datetime.date.today()
current_year = today.year
prev_month_num = today.month - 1
if len(str(prev_month_num)) == 1:
    prev_month_num_str = str("0" + str(prev_month_num))
else:
    prev_month_num_str = str(prev_month_num)
prev_month_abbr = calendar.month_abbr[prev_month_num]
prev_month_word = calendar.month_name[prev_month_num]
report_year = today.year
# modify previous month if current month is January
if prev_month_num == 0:
    prev_month_num = 12
    prev_month_abbr = calendar.month_abbr[prev_month_num]
    prev_month_word = calendar.month_name[prev_month_num]
    report_year = today.year - 1

# get paths to map out how data flows if not connected to the VPN
try:
    # faclistpath = 'P:\\PACS\\Finance\\General Info\\Finance Misc\\Facility List.xlsx'
    faclistpath = "P:\\PACS\\Finance\\Automation\\PCC Reporting\\pcc webscraping.xlsx"
    try:
        os.mkdir(userpath + '\\Documents\\PCC HUB\\')  # make directory for backup in documents folder
        shutil.copyfile(faclistpath, userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx')  # make backup file
    except FileExistsError:
        shutil.copyfile(faclistpath, userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx')  # if folder exists just copy
except FileNotFoundError:  # if VPN is not connected use the one last saved
    faclistpath = userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx'

facility_df = pd.read_excel(faclistpath, sheet_name='Automation', index_col=0)
facilities_df = pd.read_excel(faclistpath, sheet_name='Automation', index_col=0, usecols=['Common Name', 'Accountant'])


def get_time():
    # date and time info
    global report_checkboxes
    today_now = time.localtime()
    now_month = today_now.tm_mon  # month
    now_day = today_now.tm_mday  # day
    now_hour = today_now.tm_hour  # hour
    now_min = today_now.tm_min  # min
    if now_day == 15:
        if now_hour == 20:
            if now_min == 1:
                report_checkboxes = {report: IntVar() for report in reports_list}  # create dict of check_boxes
                # run the program
    print('Checked the time')


# run get_time function every 59 seconds
# timer = threading.Timer(59, get_time)
# timer.start()


def write_to_csv(filename, date, building, report, period):
    check = os.path.exists(filename)

    if check:
        with open(filename, 'a', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow([date, building, report, period])
    else:
        with open(filename, 'w', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(["Date", "Building", "Report Name", "Report Period"])

        with open(filename, 'a', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow([date, building, report, period])


# pull user's name
def getName(user=username):
    username_parse = user.split(".")
    name = ""
    for x in username_parse:
        x = x.capitalize()
        name = name + " " + x
    return name


# rename and move files
def renameDownloadedFile(newfilename, dirpath):  # renames file most recent file in downloads folder and moves it to dirpath
    global newpathtext
    try:
        newptext = newpathtext
    except NameError:
        newptext = '\\'
    time.sleep(2)
    try:
        listoffiles = glob.glob(userpath + '\\Downloads\\*')    # get a list of filesp
        latestfile = max(listoffiles, key=os.path.getctime)     # find the latest file
        extention = os.path.splitext(latestfile)[1]             # get the extension of the latest file
        if newptext != '\\':
            destfile = os.path.join(newptext, newfilename + extention)
        else:
            destfile = os.path.join(dirpath, newfilename + extention)
        try:
            shutil.move(latestfile, destfile)   # try to save file to original folder (if error with VPN)
        except:                                                             # BACKUP LOCATION IF VPN GOES DOWN
            try:                                                            # make new folder if doesn't exist
                os.mkdir(userpath + '\\Desktop\\temp reporting\\')          # temp file on desktop
                newptext = userpath + '\\Desktop\\temp reporting\\'         # temp file on desktop
                destfile = os.path.join(newptext, newfilename + extention)  # form save file path
            except FileExistsError:                                         # if folder does exist then just save
                newptext = userpath + '\\Desktop\\temp reporting\\'
                destfile = os.path.join(newptext, newfilename + extention)
            shutil.move(latestfile, destfile)                               # MOVE AND RENAME
        callback('Moved to: ' + destfile)                                   # END BACKUP LOCATION
    except:
        callback("Issue renaming/moving to " + str(dirpath))
        callback(newfilename + extention + " is in Downloads folder")


# open autofillfinancials folder
def openFinancialsFolder():  # open hidden folder on desktop that holds downloaded financials
    callback("Opening AutoFillFinancials folder")
    os.startfile(userpath + '\\Desktop\\AutoFillFinancials\\')


def convert_to_xlsx():
    listoffiles = glob.glob(userpath + '\\Downloads\\*')  # get a list of files
    latestfile = max(listoffiles, key=os.path.getctime)  # find the latest file
    extention = os.path.splitext(latestfile)[1]  # get the extension of the latest file
    # excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    # excel = win32com.client.DispatchEx("Excel.Application")
    excel = win32com.client.dynamic.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(latestfile)
    wb.SaveAs(latestfile + "x", FileFormat=51)
    wb.Close()
    excel.Application.Quit()


# get latest driver from the shared drive and add to user documents folder
def find_updated_driver():
    folder = 'P:\\PACS\\Finance\\Automation\\Chromedrivers\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        try:
            callback('Updating chromedriver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + max(file_list) + '.exe')
            callback('chromedriver updated to version '+ max(file_list))
        except:
            callback("Couldn't update chromedriver automatically")
        return max(file_list)
    else:
        callback('Could not find P:\\PACS\\Finance\\Automation\\Chromedrivers\\')


# check the current driver version on your computer
def find_current_driver():
    folder = os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\'
    file_list = []
    if os.path.isdir(folder):
        list_items = os.listdir(folder)
        for item in list_items:
            file = item.split(" ")
            if file[0] == 'chromedriver':
                file_list.append(file[1][:2])
        return max(file_list)


# start a new instance of class pcc
def startPCC():
    global PCC
    PCC = LoginPCC()


# download census report for kindred buildings
def downloadKindredReport():
    callback("Downloading weekly census report for Kindred buildings")
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
        callback('Workbook was not saved.')
    callback('Process has finished')


# initiate download of reports
def downloadIncomeStmtM2M():
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    for facname in facilities:
        if check_if_selected(facname):  # check for correct facility in PCC
            callback("Downloading income statement: " + facname)
            PCC.buildingSelect(facname)  # get next building in the list on chrome
            time.sleep(1)
            PCC.IS_M2M(str(report_year), facname)  # run reports

    callback("Income statements downloaded")


def newExcel():
    wb = xw.Book()  # create new Excel instance
    sht1 = wb.sheets["Sheet1"]  # sheet name
    wb.activate(steal_focus=True)  # focus the new instance


# run the reports
def download_reports():
    global check_status
    global PCC
    check_status = True
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    for facname in facilities:                  # loop thorugh buildings
        pcc_building = facilities[facname][2]   # pcc full name
        bu = facilities[facname][1]
        if check_if_selected(facname):          # is this building checked
            callback(facname)                   # tell user they are switching
            PCC.buildingSelect(pcc_building)    # select building in PCC
            time.sleep(1)
            for report in report_checkboxes.keys(): # get reports for this building
                if report_checkboxes[report] == 1:  # if checkbox is selected
                    if check_status:
                        if not PCC.checkSelectedBuilding(pcc_building):
                            PCC.buildingSelect(pcc_building)
                        if report == 'AP Aging':
                            callback('AP Aging')
                            PCC.ap_aging(facname)
                        if report == 'AR Aging': # uses management console
                            callback('AR Aging')
                            PCC.ar_aging(facname, bu)
                        if report == 'AR Rollforward':
                            callback('AR Rollforward')
                            PCC.ar_rollforward(facname)
                        if report == 'Cash Reciepts Journal':
                            callback('Cash Recipts Journal')
                            PCC.cash_receipts(facname)
                        if report == 'Detailed Census':
                            callback('Detailed Census')
                            PCC.census(facname)
                        if report == 'Journal Entries':
                            callback('Journal Entries')
                            PCC.journal_entries(facname)
                        if report == 'Revenue Reconciliation':
                            callback('Revenue Reconciliation')
                            PCC.revenuerec(facname)
                    else:
                        callback('There is an issue with the chromedriver')
    callback('Reports downloaded')
    PCC.teardown_method()
    del PCC


class LoginPCC:
    def __init__(self):  # create an instance of this class. Begins by logging in
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
            prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings), "plugins.always_open_pdf_externally": True}
            chrome_options.add_experimental_option('prefs', prefs)
            chrome_options.add_argument('--kiosk-printing')
            try:
                # chromedriver_autoinstaller.install()
                latestdriver = find_current_driver()
                self.driver = webdriver.Chrome(os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe', options=chrome_options)
                callback('Local chromedriver successfully initiated')
            except:
                latestdriver = find_updated_driver()
                self.driver = webdriver.Chrome(
                    os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                    options=chrome_options)
                callback('Chromedrive successfully initiated')
            try:
                self.driver.get('https://login.pointclickcare.com/home/userLogin.xhtml?ESOLGuid=40_1572368815140')
                time.sleep(3)
                try:
                    username = self.driver.find_element(By.ID, 'username')
                    username.send_keys(usernametext)
                    password = self.driver.find_element(By.ID, 'password')
                    password.send_keys(passwordtext)
                    self.driver.find_element(By.ID, 'login-button').click()
                    time.sleep(3)
                except:
                    usernamex = self.driver.find_element(By.ID, 'id-un')
                    usernamex.send_keys(usernametext)
                    passwordx = self.driver.find_element(By.ID, 'password')
                    passwordx.send_keys(passwordtext)
                    self.driver.find_element(By.ID, 'id-submit').click()
            except:
                print("There is an issue with the chrome driver")
        except:
            callback('There was an issue initiating chromedriver')
            check_status = False

    def teardown_method(self):  # exit the program (FULLY WORKING)
        try:
            self.driver.quit()
        except:
            pass

    def close_all_windows(self, firstwindow):
        original_window = firstwindow
        all_windows = self.driver.window_handles
        for window in all_windows:
            if window != original_window:
                self.driver.switch_to.window(window)
                self.driver.close()
        self.driver.switch_to.window(firstwindow)

    def buildingSelect(self, building):  # select your building (FULLY WORKING)
        self.driver.get("https://www12.pointclickcare.com/emc/home.jsp")
        self.driver.find_element(By.ID, "pccFacLink").click()
        time.sleep(1)
        try:
            self.driver.find_element(By.PARTIAL_LINK_TEXT, building).click() # select the building
            time.sleep(2)                                                    # wait
            if building in self.driver.find_element(By.ID, "pccFacLink").get_attribute("title"):
                pass
            else:
                callback('Could not find ' + building)
        except:
            callback('Could not get the proper page')

    def checkSelectedBuilding(self, building):
        try:
            if building in self.driver.find_element(By.ID, "pccFacLink").get_attribute("title"):
                return True
            else:
                return False
        except:
            self.driver.get("https://www12.pointclickcare.com/emc/home.jsp")
            if building in self.driver.find_element(By.ID, "pccFacLink").get_attribute("title"):
                return True
            else:
                return False

    def IS_M2M(self, year, facname):  # download the income statement m-to-m report (FULLY WORKING)
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get('https://www12.pointclickcare.com/glap/reports/glapreports.jsp')
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "Income Statement - System with Census - M-to-M").click()
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
                userpath + "\\Desktop\\AutoFillFinancials\\")  # rename and move file
        except:
            callback('There was an issue downloading')
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'IS M2M', str(prev_month_num) + " " + str(report_year))

    def ap_aging(self, facname):  # download AP aging report. Paste to Excel (FULLY WORKING)
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/glap/reports/glapreports.jsp")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "A/P Trial Balance - NEW").click()
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
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'a')  # highlight the entire page
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # copy the entire page
            time.sleep(2)
            self.close_all_windows(window_before)
            wb = xw.Book()  # new workbook
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
            try:
                wb.save(newpathtext + str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                wb.close()
                callback(facname + ' AP aging saved')
            except:
                try:
                    wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AP Aging\\" +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                    wb.close()
                    callback(facname + ' AP aging saved')
                except:
                    try:
                        os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                        wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                        wb.close()
                        callback(facname + ' AP aging saved')
                    except:
                        try:
                            wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                    str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                            wb.close()
                            callback(facname + ' AP aging saved')
                        except:
                            callback('Error saving AP aging')
                time.sleep(2)
        except:
            callback('Issue downloading AP Aging: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AP AGING',
                         str(prev_month_num) + " " + str(report_year))

    def ar_aging(self,facname, bu):  # pull ar aging files - (FULLY WORKING) Saves as Excel file
        try:
            iter = True
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            title = self.driver.find_element(By.ID, "pccFacLink")
            time.sleep(1)
            if title.text != "Enterprise Management Console":
                iter = False
                self.driver.get("https://www12.pointclickcare.com/emc/home.jsp")
                self.driver.find_element(By.ID, "pccFacLink").click()
                time.sleep(1)
                self.driver.find_element(By.CSS_SELECTOR, "#facTabs .pccButton").click()  # go to management console
                time.sleep(1)
                self.driver.get("https://www12.pointclickcare.com/emc/reporting.jsp?EMCmodule=P")  # go to reports
                self.driver.find_element(By.LINK_TEXT, "AR Aging").click()                      # go to ar aging report
            self.driver.find_element(By.LINK_TEXT, "select").click()                            # click facilities
            window_after = self.driver.window_handles[1]                                        # set second tab
            self.driver.switch_to.window(window_after)                                          # select the second tab
            self.driver.find_element(By.CSS_SELECTOR, "#footer > input:nth-child(2)").click()   # clear all
            self.driver.find_element(By.ID, "ESOLfacid_" + str(bu)).click()                     # select building
            self.driver.find_element(By.CSS_SELECTOR, ".pccButton:nth-child(3)").click()        # save and exit
            self.driver.switch_to.window(window_before)                                         # go back to original tab.  Facility is selected
            if not iter:
                dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")                 # select the reporting date
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
                    callback('Issue moving and renaming the file')
            except:
                callback('Issue converting excel file')
        except:
            callback('Issue downloading AR Aging: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AR AGING',
                         str(prev_month_num) + " " + str(report_year))

    def ar_rollforward(self, facname):  # download ar rollforward report.(FULLY WORKING) Paste to Excel
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/admin/reports/adminreports.jsp?ESOLtabtype=P")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "A/R Reconciliation").click()
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
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'a')  # highlight the entire page
            self.driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, 'c')  # highlight the entire page
            time.sleep(5)
            self.close_all_windows(window_before)
            wb = xw.Book()  # new workbook
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
            try:
                wb.save(newpathtext + str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                wb.close()
                callback(facname + ' AR Rollforward saved')
            except:
                try:
                    wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AR Rollforward\\" +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                    wb.close()
                    callback(facname + ' AR Rollforward saved')
                except:
                    try:
                        os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                        wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                        wb.close()
                        callback(facname + ' AR Rollforward saved')
                    except:
                        try:
                            wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                    str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                            wb.close()
                            callback(facname + ' AR Rollforward saved')
                        except:
                            callback('Error saving AR Rollforward')
                time.sleep(2)
        except:
            callback('Issue downloading AR Rollforward: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AR ROLLFORWARD',
                         str(prev_month_num) + " " + str(report_year))

    def cash_receipts(self, facname):  # prints to PDF -WORKED PERFECTLY
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/admin/reports/adminreports.jsp?ESOLtabtype=P")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "Cash Receipts Journal").click()
            self.driver.find_element(By.NAME, "ESOLdateselect_active").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = self.driver.find_element(By.NAME, "ESOLmonthSelect")
            dropdown.find_element(By.CSS_SELECTOR, "#pickdate > select:nth-child(2) > option:nth-child(" + str(prev_month_num) + ")").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            self.driver.find_element(By.NAME, "ESOLmonthSelect").click()
            dropdown = Select(self.driver.find_element(By.NAME, "ESOLyearSelect"))
            dropdown.select_by_value(str(report_year))
            self.driver.find_element(By.ID, "runButton").click()
            time.sleep(5)  # wait
            window_after = self.driver.window_handles[1]    # set second tab
            self.driver.switch_to.window(window_after)      # select the second tab
            self.driver.execute_script('window.print();')   # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Cash Receipts',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Cash Receipts\\')
        except:
            callback('Issue downloading Cash Receipts: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'CASH RECEIPTS',
                         str(prev_month_num) + " " + str(report_year))

    def census(self, facname):  # prints to PDF -WORKED PERFECTLY
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/admin/reports/adminreports.jsp?ESOLtabtype=P")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "Detailed Census").click()
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
            window_after = self.driver.window_handles[1]    # set second tab
            self.driver.switch_to.window(window_after)      # select the second tab
            self.driver.execute_script('window.print();')   # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Census',
                                 'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Census\\')
        except:
            callback('Issue downloading Census: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'CENSUS',
                         str(prev_month_num) + " " + str(report_year))

    def journal_entries(self, facname):  # prints to PDF
        try:
            time.sleep(1)
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get("https://www12.pointclickcare.com/admin/reports/adminreports.jsp?ESOLtabtype=P")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "Journal Entries").click()
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
            callback('Issue downloading Journal Entries: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'JOURNAL ENTRIES',
                         str(prev_month_num) + " " + str(report_year))

    def revenuerec(self, facname):  # prints to PDF
        try:
            window_before = self.driver.window_handles[0]  # make window tab object
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/admin/reports/adminreports.jsp?ESOLtabtype=P")
            time.sleep(1)
            self.driver.find_element(By.LINK_TEXT, "Revenue Reconciliation").click()
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
            self.driver.execute_script('window.print();')   #print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Revenue Reconciliation',
                'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Revenue Reconciliation\\')
        except:
            callback('Issue downloading Revenue Reconciliation: ' + facname)
            write_to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'REVENUE RECON',
                         str(prev_month_num) + " " + str(report_year))

    def close_ap_periods(self):  # might have issues at end of the year
        try:
            print("Closing the month of " + str(prev_month_abbr))
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get('https://www12.pointclickcare.com/glap/setup/glapsetup.jsp')
            self.driver.find_element(By.LINK_TEXT, "Fiscal Calendar Setup").click()
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
            callback('There was an issue downloading')

    def close_gl_periods(self):  # might have issues at end of the year
        try:
            # print("Closing the month of " + str(prev_month_abbr))
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get('https://www12.pointclickcare.com/glap/setup/glapsetup.jsp')
            self.driver.find_element(By.LINK_TEXT, "Fiscal Calendar Setup").click()
            self.driver.find_element(By.CSS_SELECTOR,
                                     "#expandoTabDivGL > div > table > tbody > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)").click()
            window_after = self.driver.window_handles[1]  # set second tab
            self.driver.switch_to.window(window_after)  # select the second tab
            self.driver.find_element(By.XPATH, "//*[@id=\"fiscal" + str(prev_month_num) + "c\"]").click()
            self.driver.find_element(By.CSS_SELECTOR, "#msg > input:nth-child(1)").click()  # click button weekly
            time.sleep(2)
            self.driver.switch_to.window(window_before)  # go back to original tab
        except:
            callback('There was an issue downloading')

    def kindredReport(self):  # download kindred report.  (FULLY WORKING)
        try:
            tdy = datetime.date.today()
            if tdy.weekday() == 6:
                sunday = datetime.date.today()
            else:
                sunday = (tdy + datetime.timedelta(days=(-tdy.weekday() - 1), weeks=0))  # gets previous monday
            sundaystr = str(sunday.month) + "/" + str(sunday.day) + "/" + str(sunday.year)
            # callback("Pulling report for date ending " + str(sundaystr))
            window_before = self.driver.window_handles[0]  # make window tab object
            self.driver.get("https://www12.pointclickcare.com/emc/home.jsp")
            self.driver.find_element(By.ID, "pccFacLink").click()
            time.sleep(1)
            self.driver.find_element(By.CSS_SELECTOR, "#facTabs .pccButton").click()
            # self.driver.find_element(By.XPATH, '//*[@id="facTabs"]/tbody/tr/td[2]/input]')
            time.sleep(1)
            self.driver.get("https://www12.pointclickcare.com/emc/reporting.jsp?EMCmodule=P")
            self.driver.find_element(By.LINK_TEXT, "Detailed Census Reports").click()
            self.driver.find_element(By.CSS_SELECTOR, "tr:nth-child(3) label:nth-child(1)").click()  # click button weekly
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
            self.driver.close()
            self.driver.switch_to.window(window_before)  # go back to original tab
        except:
            callback('There was an issue downloading')


# GUI SECTION *************************************************************************************************
root = Tk()  # create a GUI
root.title("Providence Group Accounting Hub v1.20.05.11")
# root.geometry("%dx%d+%d+%d" % (1200, 400, 1000, 200))
root.resizable(False, False)
# root.iconbitmap("PACS.ico")

# facilities = facility_df['Accountant'].to_dict()  # make facility dict for checkboxes (facname, accountant)
facilities = facility_df.index.to_list()
accountants = facility_df['Accountant'].to_list()
fac_number = facility_df['Business Unit'].to_list()
pcc_name = facility_df['PCC Name'].to_list()
facilities = dict(zip(facilities, zip(accountants, fac_number, pcc_name)))
accountantlist = facility_df['Accountant'].drop_duplicates().to_list()  # make list of all accountants

reports_list = ['AP Aging',
                'AR Aging',
                'AR Rollforward',
                'Cash Reciepts Journal',
                'Detailed Census',
                'Journal Entries',
                'Revenue Reconciliation']


# new window for checkboxes
def new_winF():
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Select Facility")
    newwin.resizable(False, False)
    # newwin.iconbitmap("PACS.ico")

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=3)
    scframe.grid(row=1, column=0)
    saframe.grid(row=1, column=1)
    caframe.grid(row=1, column=2)

    def get_value():  # get checkbox status and close window
        for status in check_boxes:
            check_boxes[status] = check_boxes[status].get()
        newwin.destroy()

    def select_all():
        for status in check_boxes:
            check_boxes[status].set(1)

    def clear_all():
        for status in check_boxes:
            check_boxes[status].set(0)

    global check_boxes
    check_boxes = {facility: IntVar() for facility in newfacilities}  # create dict of check_boxes
    i = 0
    r = 0
    for facility in newfacilities:  # loop to add boxes from list
        i += 1
        box = Checkbutton(boxframe, text=facility, variable=check_boxes[facility], bg=headcolor)
        if i <= 10:
            box.grid(row=i, column=r, sticky=W)
        else:
            r += 1
            i = 1
            box.grid(row=i, column=r, sticky=W)

    savebutton = Button(scframe, padx=2, pady=2, width=15, text="Save and Close", command=get_value)
    selectallbutton = Button(saframe, padx=2, pady=2, width=15, text="Select All", command=select_all)
    clearallbutton = Button(caframe, padx=2, pady=2, width=15, text="Clear All", command=clear_all)

    savebutton.grid(row=11, sticky="nsew")
    selectallbutton.grid(row=11, column=1, sticky="nsew")
    clearallbutton.grid(row=11, column=2, sticky="nsew")
    newwin.mainloop()


def select_reports_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Month End Reports")
    newwin.resizable(False, False)

    # newwin.iconbitmap("PACS.ico")
    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)


    boxframe.grid(row=0, columnspan=3)
    scframe.grid(row=1, column=0)
    saframe.grid(row=1, column=1)
    caframe.grid(row=1, column=2)

    def get_value():  # get checkbox status and close window
        global newpathtext
        newpathtext = newpath.get() + '\\'
        if newpathtext == 'temp\\':
            try:
                os.mkdir(userpath + '\\Desktop\\temp reporting\\')       # temp file on desktop
                newpathtext = userpath + '\\Desktop\\temp reporting\\'   # temp file on desktop
            except FileExistsError:
                newpathtext = userpath + '\\Desktop\\temp reporting\\'
        for status in report_checkboxes:
            report_checkboxes[status] = report_checkboxes[status].get()
        newwin.destroy()

    def select_all():
        for status in report_checkboxes:
            report_checkboxes[status].set(1)

    def clear_all():
        for status in report_checkboxes:
            report_checkboxes[status].set(0)

    global report_checkboxes
    report_checkboxes = {report: IntVar() for report in reports_list}  # create dict of check_boxes
    i = 0
    r = 0
    for report in reports_list:  # loop to add boxes from list
        i += 1
        box = Checkbutton(boxframe, text=report, variable=report_checkboxes[report], bg=headcolor)
        if i <= 10:
            box.grid(row=i, column=r, sticky=W)
        else:
            r += 1
            i = 1
            box.grid(row=i, column=r, sticky=W)

    pathlabel = Label(boxframe, text="New Folder ('temp' for desktop): ", bg=headcolor)
    pathlabel.grid(row=i+1, column=0)
    newpath = Entry(boxframe)
    newpath.grid(row=i+1, column=1, pady=5, padx=5)

    savebutton = Button(scframe, padx=2, pady=2, width=15, text="Save and Close", command=get_value)
    selectallbutton = Button(saframe, padx=2, pady=2, width=15, text="Select All", command=select_all)
    clearallbutton = Button(caframe, padx=2, pady=2, width=15, text="Clear All", command=clear_all)

    savebutton.grid(row=11, sticky="nsew")
    selectallbutton.grid(row=11, column=1, sticky="nsew")
    clearallbutton.grid(row=11, column=2, sticky="nsew")

    newwin.mainloop()


def reports_login_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("PCC Login")
    newwin.resizable(False, False)

    # newwin.iconbitmap("PACS.ico")

    def getentrytext():
        global usernametext
        global passwordtext
        usernametext = username.get()
        passwordtext = password.get()
        newwin.destroy()
        download_reports()

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)

    # login info labels
    name = getName()
    usernamelabel = Label(saframe, text="PCC Username", bg=headcolor)
    usernamelabel.grid(row=3, column=0)
    passwordlabel = Label(saframe, text="PCC Password", bg=headcolor)
    passwordlabel.grid(row=4, column=0)
    # login info entry
    username = Entry(saframe)
    username.insert(0, "pghc." + name[1].lower() + name.split(" ")[2].lower())
    username.grid(row=3, column=1)
    password = Entry(saframe, show="*")
    password.grid(row=4, column=1)

    runbutton = Button(scframe, padx=2, pady=2, width=15, text="Run", command=getentrytext)
    runbutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


def kindred_login_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("PCC Login")
    newwin.resizable(False, False)

    # newwin.iconbitmap("PACS.ico")

    def getentrytext():
        global usernametext
        global passwordtext
        usernametext = username.get()
        passwordtext = password.get()
        newwin.destroy()
        downloadKindredReport()

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)

    # login info labels
    name = getName()
    usernamelabel = Label(saframe, text="PCC Username", bg=headcolor)
    usernamelabel.grid(row=3, column=0)
    passwordlabel = Label(saframe, text="PCC Password", bg=headcolor)
    passwordlabel.grid(row=4, column=0)
    # login info entry
    username = Entry(saframe)
    username.insert(0, "pghc." + name[1].lower() + name.split(" ")[2].lower())
    username.grid(row=3, column=1)
    password = Entry(saframe, show="*")
    password.grid(row=4, column=1)

    runbutton = Button(scframe, padx=2, pady=2, width=15, text="Run", command=getentrytext)
    runbutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


def M2M_login_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("PCC Login")
    newwin.resizable(False, False)

    # newwin.iconbitmap("PACS.ico")

    def getentrytext():
        global usernametext
        global passwordtext
        usernametext = username.get()
        passwordtext = password.get()
        newwin.destroy()
        downloadIncomeStmtM2M()

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)

    # login info labels
    name = getName()
    usernamelabel = Label(saframe, text="PCC Username", bg=headcolor)
    usernamelabel.grid(row=3, column=0)
    passwordlabel = Label(saframe, text="PCC Password", bg=headcolor)
    passwordlabel.grid(row=4, column=0)
    # login info entry
    username = Entry(saframe)
    username.insert(0, "pghc." + name[1].lower() + name.split(" ")[2].lower())
    username.grid(row=3, column=1)
    password = Entry(saframe, show="*")
    password.grid(row=4, column=1)

    runbutton = Button(scframe, padx=2, pady=2, width=15, text="Run", command=getentrytext)
    runbutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


def update_date_win():  # new window definition
    newwin = Toplevel(root, bg=headcolor)
    newwin.title("Post check run batches")
    newwin.resizable(False, False)

    # newwin.iconbitmap("PACS.ico")

    def getentrytext():
        global newmonthtext
        global newyeartext
        newmonthtext = newmonth.get()
        newyeartext = newyear.get()
        try:
            if int(newmonthtext) <= 12 and 3 < len(newyeartext) < 5:
                newwin.destroy()
                update_date(newmonthtext, newyeartext)
        except ValueError:
            callback('No date input, please try a different date.')

    boxframe = Frame(newwin, bg=headcolor, pady=10, bd=10)
    scframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    saframe = Frame(newwin, bg=headcolor, pady=4, bd=10)
    caframe = Frame(newwin, bg=headcolor, pady=4, bd=10)

    # newwin.grid_rowconfigure(0, weight=1)
    # newwin.grid_columnconfigure(0, weight=1)

    boxframe.grid(row=0, columnspan=2)
    scframe.grid(row=2, column=0)
    saframe.grid(row=1, column=0)
    welcomelabel.grid(row=0, column=0)
    welcomelabel.config(font=22)

    # login info labels
    monthlabel = Label(saframe, text="Month: ", bg=headcolor)
    monthlabel.grid(row=1, column=0)
    yearlabel = Label(saframe, text="Year: ", bg=headcolor)
    yearlabel.grid(row=2, column=0)

    # login info entry
    newmonth = Entry(saframe)
    newmonth.grid(row=1, column=1)
    newyear = Entry(saframe)
    newyear.grid(row=2, column=1)

    savebutton = Button(scframe, padx=2, pady=2, width=15, text="Save", command=getentrytext)
    savebutton.grid(row=11, sticky="nsew")

    newwin.mainloop()


# let user change date
def update_date(monthinput='', yearinput=''):
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
    closingdate.set("Selected Month: " + prev_month_abbr + " " + str(report_year))
    callback('Reporting date changed to ' + prev_month_abbr + ' ' + str(report_year))


# print to GUI box
def callback(message):  # update the statusbox gui
    s = str(datetime.datetime.now().strftime("%H:%M:%S")) + ">>" + str(message) + "\n"
    statusbox.insert(END, s)
    statusbox.see(END)
    statusbox.update()


# create dropdown menu for accountants
def boxtext(new_value):
    item = dropbox.get()
    global newfacilities
    newfacilities = []
    if item == "All":
        for building in facilities:
            newfacilities.append(building)
    else:
        for building in facilities:
            if facilities[building][0] == item:
                newfacilities.append(building)


# check if the building checkbox is selected
def check_if_selected(building):
    try:
        if building in check_boxes.keys():
            if check_boxes[building] == 0:
                # callback(building + " is not selected.  Going to next facility")
                return False
            else:
                # callback(building)
                return True
    except:
        # callback(building + " is not selected.  Trying next facility")
        pass


# def show_selected_reports():
#     global PCC
#     try:
#         for report in report_checkboxes.keys():
#             if report_checkboxes[report] == 1:
#                 download_reports(report)
#     except NameError:
#         callback('No reports selected')
#     callback('Reports downloaded')
#     PCC.teardown_method()
#     del PCC


headcolor = "#d7eef5"
framecolor = "#d7eef5"
footcolor = "#d7eef5"
statusboxcolor = "#f7f7f7"

# create the frames
headframe = Frame(root, width=400, height=300, bg=headcolor, pady=10, bd=10)
downloadsframe = Frame(root, width=400, height=400, bg=framecolor, padx=10, bd=10)
otherframe = Frame(root, width=400, height=400, bg=framecolor, padx=10, bd=10)
footframe = Frame(root, width=400, height=100, bg=footcolor, pady=10, bd=20)

# layout all of the main containters
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

headframe.grid(row=0, sticky="nsew", columnspan=10)
downloadsframe.grid(row=1, column=0, sticky="nsew")
otherframe.grid(row=1, column=1, sticky="nsew")
footframe.grid(row=2, sticky="nsew", columnspan=10)

statusbox = Text(footframe, height=20, width=70, bg=statusboxcolor)
statusbox.grid(row=1, column=0, columnspan=10)

# create the labels - headframe
welcomelabel = Label(headframe, text="Welcome" + getName(), bg=headcolor)
welcomelabel.grid(row=0, column=0)
welcomelabel.config(font=88)
closingdate = StringVar()
closingmonthlabel = Label(headframe, textvariable=closingdate, bg=headcolor)
closingdate.set("Selected Month: " + prev_month_abbr + " " + str(report_year))
closingmonthlabel.grid(row=2, column=0, sticky="nsew")

# dropdown box for accountant names
data = accountantlist                           # create list
data = [x for x in data if str(x) != 'nan']     # get rid of 'nan' values
data.reverse()                                  # reverse list to add 'All' to the front
data.append('All')                              # add 'All' to end
data.reverse()                                  # revert list back
dropbox = StringVar()
dropbox.set('All')                              # set dropdown default to 'All'
boxtext("All")                                  # run function to setup the correct check boxes based on dropbox all
dropdownbox = OptionMenu(headframe, dropbox, *data, command=boxtext)
dropdownbox.grid(row=3, column=0)               # add to GUI

# create the buttons
# downloads frame buttons - middle frames
selectedreportsbutton = Button(downloadsframe, text="Download Selected Reports", padx=5, pady=5, width=35, command=reports_login_win)
incomestatementbutton = Button(downloadsframe, text="Download Income Statements M-to-M", padx=5, pady=5, width=35, command=M2M_login_win)
kindredreportbutton = Button(downloadsframe, text="Download Kindred Report", padx=5, pady=5, width=35, command=kindred_login_win)

# add the buttons
selectedreportsbutton.grid(row=0, padx=5, pady=5, sticky="nsew")
incomestatementbutton.grid(row=1, padx=5, pady=5, sticky="nsew")
kindredreportbutton.grid(row=2, padx=5, pady=5, sticky="nsew")

# other frame buttons
choosefacbutton = Button(otherframe, text="Select Facilities", padx=5, pady=5, width=35, command=new_winF)  # command linked
showselectedfac = Button(otherframe, text="Select Reports", padx=5, pady=5, width=35, command=select_reports_win)  # command linked
changedatesbutton = Button(otherframe, text="Change Date", padx=5, pady=5, width=35, command=update_date_win)
facfolderbutoon = Button(otherframe, text="Open Facility Folders", padx=5, pady=5, width=35, command='')
financialfolderbutton = Button(otherframe, text="Open AutoFillFinancials Folder", padx=5, pady=5, width=35, command=openFinancialsFolder)
closeapperiodsbutton = Button(otherframe, text="Close AP Period", padx=5, pady=5, width=35, command='')
closeglperiodsbutton = Button(otherframe, text="Close GL Period", padx=5, pady=5, width=35, command='')
copyfinancialsbutton = Button(otherframe, text="Copy Over Financials (not working)", padx=5, pady=5, width=35, command='')


# add the buttons
choosefacbutton.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
showselectedfac.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
changedatesbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")
# closeapperiodsbutton.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")
# closeglperiodsbutton.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")
# facfolderbutoon.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")
# financialfolderbutton.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")
# copyfinancialsbutton.grid(row=7, column=0, padx=5, pady=5, sticky="nsew")


statuslabel = Label(footframe, text="Status Box:", bg=footcolor)
statuslabel.grid(row=0, column=0, sticky="nsew")

root.mainloop()

# def say_hello(systray):
#     print("Hello, World!")


# menu_options = (("Say Hello", None, say_hello),)
# systray = SysTrayIcon("icon.ico", "PACS Reporting", menu_options)
# systray.start()


