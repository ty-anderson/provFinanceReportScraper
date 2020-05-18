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
import csv
import win32com
from infi.systray import SysTrayIcon
import multitimer
from PIL import Image, ImageDraw,ImageFont

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
    today_now = time.localtime()
    now_month = today_now.tm_mon  # month
    now_day = today_now.tm_mday  # day
    now_hour = today_now.tm_hour  # hour
    now_min = today_now.tm_min  # min
    if now_day == 15:
        if now_hour == 20: # 8:00 pm autorun
            if now_min == 1:
                download_reports()
    print('Not time')


def to_csv(building='', report='', period='',date=str(datetime.date.today()) + ' ' + str(time.localtime().tm_hour)
                    + ':' + str(time.localtime().tm_min), filename=userpath + '\\Desktop\\PyReport.csv'):
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


def to_text(message):
    s = str(datetime.datetime.now().strftime("%H:%M:%S")) + ">>" + str(message) + "\n"
    with open(userpath + '\\Desktop\\PyReport.csv','a') as file:
        file.write(s)
        file.close()

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
        to_csv('Moved to: ' + destfile)                                   # END BACKUP LOCATION
    except:
        to_csv("Issue renaming/moving to " + str(dirpath))
        to_csv(newfilename + extention + " is in Downloads folder")


# open autofillfinancials folder
def openFinancialsFolder():  # open hidden folder on desktop that holds downloaded financials
    to_csv("Opening AutoFillFinancials folder")
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
            to_csv('Updating chromedriver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + max(file_list) + '.exe')
            to_csv('chromedriver updated to version '+ max(file_list))
        except:
            to_csv("Couldn't update chromedriver automatically")
        return max(file_list)
    else:
        to_csv('Could not find P:\\PACS\\Finance\\Automation\\Chromedrivers\\')


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
    to_csv("Downloading weekly census report for Kindred buildings")
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
        to_csv('Workbook was not saved.')
    to_csv('Process has finished')


# run the reports
def download_reports():
    global check_status
    global PCC
    download_num = len(facilities) * len(reports_list)
    update_icon(download_num)
    check_status = True
    try:
        PCC  # check if an instance already exists
    except NameError:  # if not
        startPCC()  # create one
    for facname in facilities:                  # loop through all buildings
        pcc_building = facilities[facname][2]   # pcc full name
        bu = facilities[facname][1]             # business unit
        PCC.buildingSelect(pcc_building)        # select building in PCC
        time.sleep(1)
        for report in reports_list:             # get reports for this building
            if check_status:
                if not PCC.checkSelectedBuilding(pcc_building): # verify that we have the right building selected
                    PCC.buildingSelect(pcc_building)            # if not then get the correct building
                if report == 'AP Aging':
                    to_csv(facname, 'AP Aging')
                    PCC.ap_aging(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'AR Aging': # uses management console
                    to_csv('AR Aging')
                    PCC.ar_aging(facname, bu)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'AR Rollforward':
                    to_csv('AR Rollforward')
                    PCC.ar_rollforward(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'Cash Reciepts Journal':
                    to_csv('Cash Recipts Journal')
                    PCC.cash_receipts(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'Detailed Census':
                    to_csv('Detailed Census')
                    PCC.census(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'Journal Entries':
                    to_csv('Journal Entries')
                    PCC.journal_entries(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
                if report == 'Revenue Reconciliation':
                    to_csv('Revenue Reconciliation')
                    PCC.revenuerec(facname)
                    download_num = download_num - 1
                    update_icon(download_num)
            else:
                to_csv('There is an issue with the chromedriver')
    to_csv('Reports downloaded')
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
                to_csv('Local chromedriver successfully initiated')
            except:
                latestdriver = find_updated_driver()
                self.driver = webdriver.Chrome(
                    os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                    options=chrome_options)
                to_csv('Chromedrive successfully initiated')
            try:
                self.driver.get('https://login.pointclickcare.com/home/userLogin.xhtml?ESOLGuid=40_1572368815140')
                time.sleep(3)
                f = open("info.txt", "r")
                u = f.readline()
                p = f.readline()
                try:
                    username = self.driver.find_element(By.ID, 'username')
                    username.send_keys(u)
                    password = self.driver.find_element(By.ID, 'password')
                    password.send_keys(p)
                    self.driver.find_element(By.ID, 'login-button').click()
                    time.sleep(3)
                except:
                    usernamex = self.driver.find_element(By.ID, 'id-un')
                    usernamex.send_keys(u)
                    passwordx = self.driver.find_element(By.ID, 'password')
                    passwordx.send_keys(p)
                    self.driver.find_element(By.ID, 'id-submit').click()
            except:
                print("There is an issue with the chrome driver")
        except:
            to_csv('There was an issue initiating chromedriver')
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
                to_csv('Could not find ' + building)
        except:
            to_csv('Could not get the proper page')

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
                userpath + "\\Desktop\\")  # rename and move file
        except:
            to_csv('There was an issue downloading')
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'IS M2M', str(prev_month_num) + " " + str(report_year))

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
                to_csv(facname + ' AP aging saved')
            except:
                try:
                    wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AP Aging\\" +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                    wb.close()
                    to_csv(facname + ' AP aging saved')
                except:
                    try:
                        os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                        wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                        wb.close()
                        to_csv(facname + ' AP aging saved')
                    except:
                        try:
                            wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                    str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AP Aging.xlsx')
                            wb.close()
                            to_csv(facname + ' AP aging saved')
                        except:
                            to_csv('Error saving AP aging')
                time.sleep(2)
        except:
            to_csv('Issue downloading AP Aging: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AP AGING',
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
            window_after = self.driver.window_handles[1]                                        # set second window
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
                    to_csv('Issue moving and renaming the file')
            except:
                to_csv('Issue converting excel file')
        except:
            to_csv('Issue downloading AR Aging: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AR AGING',
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
                to_csv(facname + ' AR Rollforward saved')
            except:
                try:
                    wb.save("P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\AR Rollforward\\" +
                            str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                    wb.close()
                    to_csv(facname + ' AR Rollforward saved')
                except:
                    try:
                        os.mkdir(userpath + '\\Desktop\\temp reporting\\')
                        wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                        wb.close()
                        to_csv(facname + ' AR Rollforward saved')
                    except:
                        try:
                            wb.save(userpath + '\\Desktop\\temp reporting\\' +
                                    str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' AR Rollforward.xlsx')
                            wb.close()
                            to_csv(facname + ' AR Rollforward saved')
                        except:
                            to_csv('Error saving AR Rollforward')
                time.sleep(2)
        except:
            to_csv('Issue downloading AR Rollforward: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'AR ROLLFORWARD',
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
            to_csv('Issue downloading Cash Receipts: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'CASH RECEIPTS',
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
            to_csv('Issue downloading Census: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'CENSUS',
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
            to_csv('Issue downloading Journal Entries: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'JOURNAL ENTRIES',
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
            to_csv('Issue downloading Revenue Reconciliation: ' + facname)
            to_csv(userpath + '\\Desktop\\Reports log.csv', str(datetime.date.today()), facname, 'REVENUE RECON',
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
            to_csv('There was an issue downloading')

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
            to_csv('There was an issue downloading')

    def kindredReport(self):  # download kindred report.  (FULLY WORKING)
        try:
            tdy = datetime.date.today()
            if tdy.weekday() == 6:
                sunday = datetime.date.today()
            else:
                sunday = (tdy + datetime.timedelta(days=(-tdy.weekday() - 1), weeks=0))  # gets previous monday
            sundaystr = str(sunday.month) + "/" + str(sunday.day) + "/" + str(sunday.year)
            # to_csv("Pulling report for date ending " + str(sundaystr))
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
            to_csv('There was an issue downloading')


def update_icon(number=0):
    # create image
    image = 'pil_icon.ico'
    img = Image.new('RGBA', (50, 50), color=(255, 255, 255, 90))  # color background =  white  with transparency
    d = ImageDraw.Draw(img)
    d.rectangle([(0, 40), (50, 50)], fill=(39, 112, 229), outline=None)  # color = blue
    # add text to the image
    font_type = ImageFont.truetype("arial.ttf", 30)
    a = number
    # b = n * 20
    d.text((0, 0), f"{a}", fill=(255, 255, 0), font=font_type)
    img.save(image)
    systray.update(icon=image)


def tray_download_reports(systray):
    download_reports()


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


# run get_time function every 59 seconds
monthendtimer = multitimer.MultiTimer(interval=50, function=get_time)
monthendtimer.start()

# kindredtimer = multitimer.MultiTimer(interval=(60*60), function=get_time)

# create image
image = 'pil_icon.ico'
img = Image.new('RGBA', (50, 50), color=(255, 255, 255, 90))  # color background =  white  with transparency
d = ImageDraw.Draw(img)
d.rectangle([(0, 40), (50, 50)], fill=(39, 112, 229), outline=None)  # color = blue
# add text to the image
font_type = ImageFont.truetype("arial.ttf", 30)
d.text((0, 0), f"{0}", fill=(255, 255, 0), font=font_type)
img.save(image)

menu_options = (("Run reports (auto on 15th @8pm)", None, tray_download_reports),)
systray = SysTrayIcon(image, "PACS Reporting", menu_options)
systray.start()
