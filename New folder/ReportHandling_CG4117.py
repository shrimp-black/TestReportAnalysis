import os
import re
import shutil
import configparser
import win32com.client as win32
import time
import zipfile
from bs4 import BeautifulSoup
import csv

class myconf(configparser.ConfigParser):
    def __init__(self, default=None):
        configparser.ConfigParser.__init__(self)
    def optionxform(self, optionstr):
        return optionstr

currdir = os.getcwd()
resultdic = []

def ReadINI():
    infolist = {}

    # get the current file's name and split the name into two parts(filename and extension name)
    namesplits = os.path.splitext(os.path.basename(__file__))

    # get the INI file name
    name_ini = namesplits[0] + '.ini'

    # read out the name for the CG test procedure
    config = myconf()
    config.read(name_ini)
    spreadsheet = config.get("CG4117", "Filename")

    # get general info
    for option in config.options("Info"):
        infolist[option] = config.get('Info', option)

    clearobjects = config.get("Additional", "ClearObjects")

    return spreadsheet, infolist, clearobjects

def SaveLogTofolder(foldername, filepath, report):
    folderpath = currdir + '\\' + foldername
    reportpath = folderpath + '\\' + report
    if(os.path.exists(folderpath)):
        if os.path.isfile(reportpath):
            os.remove(reportpath)
            shutil.copy(filepath, folderpath)
            os.remove(filepath)
        else:
            shutil.copy(filepath, folderpath)
            os.remove(filepath)

    else:
        os.mkdir(folderpath)
        shutil.copy(filepath, folderpath)
        os.remove(filepath)


def AttachTestReports():
    tabinfo = set()
    test_procedure_name, infolist, clearobjects = ReadINI()
    # print(infolist)
    test_procedure_path = currdir + '\\' + test_procedure_name
    excel_app = win32.Dispatch("Excel.Application")
    # excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    wb = excel_app.Workbooks.Open(test_procedure_path, UpdateLinks = False)
    sheetnumber = {'7.Virtual Grps & Symmetric Keys': 7, '8. Authenticated Msg Structure': 8, '9.MSG Anti-Replay Counter' : 9,
                   '11. MSG Authentication Cal Tbl': 10, '12. Adding Authentication MSG': 11, '13. Verifying Authenticity MSG': 12,
                   '14. MSG Authentication Key Prov': 13, '14.1 KP Stress Testing': 14, '17 Development only Calibration' : 15}
    # Clear the attachments in case the filled test reports used
    if clearobjects == '1':
        for tab in sheetnumber:
            ws = excel_app.Worksheets(sheetnumber[tab])
            ws.Activate()
            Embeded_obj = ws.OLEObjects()
            # Embeded_obj.Item(Embeded_obj.Count).Delete
            Embeded_obj.delete

    for report in os.listdir(currdir):
        filepath = currdir + '\\' + report
        if('.zip' in report):
            filename = report.split('.')[0]
            temp = filename.split('_')
            testcase = temp[0] + '_' + temp[1]
            testnum = temp[2]
            for tab in sheetnumber:
                ws = excel_app.Worksheets(sheetnumber[tab])
                ws.Activate()
                findcell = excel_app.ActiveSheet.UsedRange.Find(testcase)
                if(findcell != None):
                    # fill in the general info for the tabs which have tests covered
                    if ((testcase == 'CYS-VMA13af00_368') & (tab == '14. MSG Authentication Key Prov')):
                        continue
                    if tab not in tabinfo:
                        tabinfo.add(tab)
                        for info in infolist:
                            findinfo = excel_app.ActiveSheet.UsedRange.Find(info + ':')
                            ws.Cells(findinfo.Row, 4).Value = infolist[info]

                    Embeded_obj = ws.OLEObjects()
                    #ws.Range(ws.Cells(findcell.Row, 1), ws.Cells(600, 1))
                    findnum = ws.Range("A"+str(findcell.Row)+":"+"A600").Find(temp[2])
                    #time.sleep(2)
                    # findnum = ws.UsedRange.Find(temp[2])
                    #findnum = excel_app.ActiveSheet.UsedRange.Find(temp[2])
                    obj = Embeded_obj.Add(ClassType=None, Filename=filepath, Link=False, DisplayAsIcon=True, Width=50, Height= 50)
                    obj.Left = ws.Cells(findnum.Row, 7).left
                    obj.Top = ws.Cells(findnum.Row, 7).Top
                    a = findnum.Row
                    ws.Rows(a).RowHeight = 50
                    ws.Columns(7).ColumnWidth = 30
                    time.sleep(3)
                    SaveLogTofolder(tab, filepath, report)
                    break
                else:
                    continue
    wb.Save()
    excel_app.Quit()

def CreateReviewReport(ReviewItem):
    headers = ['Tab', 'Test Case', 'Name_Zipfile', 'Test Result', 'Test Time']
    if os._exists('CG4117 Review Result.csv'):
        os.remove('CG4117 Review Result.csv')
    with open('CG4117 Review Result.csv','w',newline='') as csv_file:
        writer = csv.DictWriter(csv_file,fieldnames=headers)
        writer.writeheader()
        for dict in ReviewItem:
            writer.writerow(dict)
        reader = csv.reader(csv_file)
        for row in reader:


def AnalyseHTMLreport():
    # Recurse the folders and find all zip files
    file = []
    filenames = os.listdir(currdir)

    for filename in filenames:
        if os.path.isdir(filename):
            subfolders = os.listdir(filename)
            for subfolder in subfolders:
                if os.path.splitext(subfolder)[1] == '.zip':
                    z = zipfile.ZipFile(currdir + '\\' + filename + '\\' + subfolder, 'r')
# Extract all the zip files and find the Frame_Report.html in order to read the test info
                    for file in z.namelist():
                        z.extract(file, "temp/")
                        # print(file)
                        newfile = "temp/" + file
                        # first_file_name = z.namelist()[0]
                        if ('Frame_Report.html' in newfile):
                            # need to open the file in the beautifulsoup otherwise unable to read the contents in the html
                            Soup = BeautifulSoup(open(newfile), 'lxml')

                            title = Soup.select('title')
                            # print(title[0].get_text())
                            testcase = title[0].get_text()

                            Testresult = Soup.find('table', attrs={'class': 'OverallResultTable'})
                            # print(Testresult.find('td').get_text())
                            testresult = Testresult.find('td').get_text()

                            defaulttable = Soup.find('table', attrs={'class': 'DefaultTable'})
                            td_list = defaulttable.find_all('td')
                            test_begin = td_list[1].text
                            test_end = td_list[4].text
                            # print(test_begin)
                            # print(test_en)
                            testtime = test_begin + test_end
                            dic = {'Tab': filename,  'Test Case': testcase, 'Name_Zipfile': subfolder, 'Test Result': testresult,
                                   'Test Time': testtime}
                            # info = [filename, testcase,testresult,testtime]
                            resultdic.append(dic)
                    z.close()
if __name__ == "__main__":
    AttachTestReports()
    AnalyseHTMLreport()
    CreateReviewReport(resultdic)
