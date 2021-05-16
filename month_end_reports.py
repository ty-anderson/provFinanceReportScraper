from dearpygui import core, simple
import pandas as pd
import shutil
import os
import datetime
import calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import glob
import json
import xlwings as xw
import pyautogui
import win32com
import pyperclip


# clear the gen_py folder that is causing issues with the xlsx conversion with win32com
try:
    shutil.rmtree(win32com.__gen_path__[:-4])
except:
    pass

global newpathtext
global PCC


def deleteDownloads():
    """Deletes everything in downloads folder"""
    filelist = glob.glob(userpath + '\\Downloads\\*')
    try:
        for f in filelist:
            os.remove(f)
    except:
        pass


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


def renameDownloadedFile(newfilename, dirpath=''):
    """Renames most recent file in downloads folder and moves it to dirpath"""
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
        if dirpath == '':
            os.rename(latestfile,userpath + '\\Downloads\\' + newfilename)
        else:
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
            update_status('Moved to: ' + destfile)                                   # END BACKUP LOCATION
    except:
        update_status("Issue renaming/moving to " + str(dirpath))
        update_status(newfilename + " is in Downloads folder")


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
            update_status('Updating chromedriver to newer version')
            shutil.copyfile(folder + 'chromedriver ' + max(file_list) + '.exe',
                            os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + max(file_list) + '.exe')
            update_status('chromedriver updated to version ' + max(file_list))
        except:
            update_status("Couldn't update chromedriver automatically")
        return max(file_list)
    else:
        update_status('Could not find P:\\PACS\\Finance\\Automation\\Chromedrivers\\')


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


# run the reports
def download_reports(facilitylist, reportlist):
    """Download month end close reports"""
    global PCC
    deleteDownloads()
    if reportlist:
        try:
            PCC  # check if an instance already exists
        except:  # if not
            startPCC()  # create one
        for facname in facilitylist:  # LOOP BUILDING LIST
            if facname == facilitylist:  # IS BUILDING CHECHED
                bu = str(facilitylist[facname][1])  # GET BU
                if len(bu) < 2:
                    bu = str(0) + bu
                if PCC.buildingSelect(bu):
                    time.sleep(1)
                    for report in reportlist:
                        if report == 'AP Aging':
                            PCC.ap_aging(facname)
                        if report == 'AR Aging':            # USES MGMT CONSOLE
                            bu = facilitylist[facname][1]     # TO SELECT BUILDING IN AR REPORT
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
            try:
                # chromedriver_autoinstaller.install()
                self.driver = webdriver.Chrome(options=chrome_options)
            except:
                try:
                    latestdriver = find_current_driver()
                    self.driver = webdriver.Chrome(
                        os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                        options=chrome_options)
                    print('Local chromedriver successfully initiated')
                except:
                    latestdriver = find_updated_driver()
                    self.driver = webdriver.Chrome(
                        os.environ['USERPROFILE'] + '\\Documents\\PCC HUB\\chromedriver ' + str(latestdriver) + '.exe',
                        options=chrome_options)
                    print('Chromedrive successfully initiated')
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
            current_fac = self.driver.find_element(By.NAME, "current_fac_id").get_attribute("value")
            if str(current_fac) != bu:
                try:
                    self.driver.find_element(By.ID, "pccFacLink").click()
                    time.sleep(1)
                    building_list = self.driver.find_element(By.ID, "optionList")
                    building_list.find_element(By.PARTIAL_LINK_TEXT, bu).click()
                    return True
                except:
                    print("Could not locate " + bu + " in PCC")
                    return False
            else:
                return True
        except:
            print("Could not find the building dropdown menu")
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
            print('There was an issue downloading')

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
                pyperclip.copy("Fiscal period has not been setup")
                wb = xw.Book()  # new workbook
                app = xw.apps.active
                time.sleep(2)
                wb.activate(steal_focus=True)  # focus the new instance
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'v')  # paste
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
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
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
                time.sleep(2)
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
            time.sleep(2)
            wb.activate(steal_focus=True)  # focus the new instance
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')  # paste
            time.sleep(2)  # wait to load
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
            self.driver.execute_script('window.print();')  # print to PDF
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
            time.sleep(10)  # wait
            self.close_all_windows(window_before)
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
            self.driver.execute_script('window.print();')  # print to PDF
            self.close_all_windows(window_before)
            renameDownloadedFile(
                str(report_year) + ' ' + prev_month_num_str + ' ' + facname + ' Revenue Reconciliation',
                'P:\\PACS\\Finance\\Month End Close\\All - Month End Reporting\\Revenue Reconciliation\\')
        except:
            print('Issue downloading Revenue Reconciliation: ' + facname)


username = os.environ['USERNAME']
userpath = os.environ['USERPROFILE']

"""Setup date info"""
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


# faclistpath = userpath + '\\Documents\\PCC HUB\\pcc webscraping.xlsx'
faclistpath = r"P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx"
building_df = pd.read_excel(faclistpath, sheet_name='Automation', usecols=['Common Name', 'Business Unit'])
buildings = building_df['Common Name']
reports_list = ['AP Aging',
                'AR Aging',
                'AR Rollforward',
                'Cash Reciepts Journal',
                'Detailed Census',
                'Journal Entries',
                'Revenue Reconciliation']



def runAll():
    month = core.get_value("Month##monthendreports")
    year = core.get_value("Year##monthendreports")
    counter = 0
    wb_ref = r"P:\PACS\Finance\Automation\PCC Reporting\pcc webscraping.xlsx"
    wb = pd.read_excel(wb_ref, sheet_name='Automation', usecols=['Common Name', 'Business Unit'])
    reports_path = [r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AP Aging',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Aging',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\AR Rollforward',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Cash Receipts',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Census',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Journal Entries',
                    r'P:\PACS\Finance\Month End Close\All - Month End Reporting\Revenue Reconciliation']
    report_names = ['AP Aging.xlsx', 'AR Aging.xlsx', 'AR Rollforward.xlsx', 'Cash Receipts.pdf',
                    'Census.pdf', 'Journal Entries.pdf', 'Revenue Reconciliation.pdf']
    print("Searching monthly reports")
    i = 0
    for path in reports_path:
        for facility in wb['Common Name']:
            file_name = path + '\\' + str(year) + ' ' + str(month) + ' ' + facility + ' ' + \
                        report_names[i]
            if not os.path.exists(file_name):
                counter = counter + 1
                print(file_name + ' missing.  Downloading now')
                rpt = report_names[i].split('.')
                rpt = [rpt[0]]
                download_reports(facility, rpt)
            else:
                size = os.path.getsize(file_name)
                kb = size/(1024)
                if kb < 10:
                    check_df = pd.read_excel(file_name)
                    check_value = check_df.iloc[0]['A']
                    if not check_value == "":
                        print(file_name + " might be empty.  Please check.")
        i += 1
    if counter == 0:
        print("Reports have all been downloaded")



"""Gui section"""


def select_all_checkboxes(sender, data):
    for building in buildings:
        core.set_value(building, True)


def clear_all_checkboxes(sender, data):
    for building in buildings:
        core.set_value(building, False)


def on_window_close(sender, data):
    print(sender)
    core.delete_item(sender)


def run_month_end_reports(sender, data):
    item_list = []
    building_list = []
    core.delete_item('Month End')
    for item in reports_list:  # LIST OF REPORTS SELECTED
        if core.get_value(item):
            item_list.append(item)
    for building in buildings:
        if core.get_value(building):
            building_list.append(building)
    update_status('Running month end reports')
    download_reports(building_list, item_list)


def update_status(message):
    print(str(message))
    core.set_value('##status box', core.get_value('##status box') + str(message) + '\n')


with simple.window('Month End', width=500, height=500, on_close=on_window_close):
    core.add_text("Select buildings")
    for building in buildings:
        core.add_checkbox(building)
    core.add_spacing()
    core.add_button("Select All", callback=select_all_checkboxes)
    core.add_same_line()
    core.add_button("Clear All", callback=clear_all_checkboxes)
    core.add_same_line()
    core.add_button("Done", callback=run_month_end_reports)
    core.add_spacing()
    core.add_separator()
    core.add_text("Select reports")
    for report in reports_list:
        core.add_checkbox(report)
    core.add_spacing()
    core.add_separator()
    core.add_spacing()
    core.add_text("Enter the month and year")
    core.add_input_text("Month##monthendreports", width=40, decimal=True, default_value=prev_month_num_str)
    core.add_input_text("Year##monthendreports", width=40, decimal=True, default_value=str(report_year))
    core.add_spacing()
    core.add_separator()
    core.add_spacing()
    core.add_button("RUN ALL", callback=runAll)


core.start_dearpygui(primary_window="Month End")