import selenium as se
from selenium import webdriver

import time
import datetime
import json

from openpyxl import Workbook
from openpyxl import load_workbook

#Read excel ID value
def workWithXl(filen):
    wb = load_workbook(filename=filen)
    sheets = wb['Sheet1']
    #print (sheets['A3'].value)
    for row in sheets.rows:
        print (row[0].value)
        row[11].value = 'Done'
    wb.save(filename=filen)


def GetCmdDict(filename):
    cmdStrs = {}

    f = open(filename,"r")
    for row in f:
        #ignore the comment line
        if str(row)[0] == '#':
            #print ('comment {0}'.format(row))
            continue
        elif str(row)[0].lower() == 'r':
            row = row.split("->")
            ky = row[0]
            val = row[1]
            cmdStrs[ky]=val

    f.close()
    return cmdStrs

if __name__ == "__main__":
    tnow = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    print ('Start at {0}'.format(tnow))
    #workWithXl('POS.xlsx')
    #open the URL
    # loginCmds = GetCmdDict("Login.prs")
    # robotCmds = GetCmdDict("Process.prs")
    # brows = webdriver.Chrome()
    # brows.get(loginCmds['URL'])
    # time.sleep(2)

    #Action by the Process.prs
    #Robot Process 1:Login
    #robot_switch (click); robot_emailid (sendkeys); robot_password(sendkeys);robot_login(click)
    # brows.find_element_by_xpath(xpath=robotCmds['robot_login']).click()
    # brows.find_element_by_xpath(xpath=robotCmds['robot_sw']).click()
    # brows.find_element_by_xpath(xpath=robotCmds['robot_usr']).send_keys(loginCmds['username'])
    # brows.find_element_by_xpath(xpath=robotCmds['robot_pwd']).send_keys(loginCmds['password'])
    # brows.find_element_by_xpath(xpath=robotCmds['robot_chsb']).click()
    # brows.find_element_by_xpath(xpath=robotCmds['robot_logbtn']).click()

    # Read ID Look
    wb = load_workbook(filename='ceb.xlsx')

    #Process POS
    sheetPOS = wb['POS']
    for row in sheetPOS.rows:
        # Robot Process 2:Process Delete Action 1 - user terminators management
        row[-1].value = 'Del action 1 done'
        # Robot Process 3:Process Delete Action 2 - user information management
        row[-2].value = 'Del action 2 done'

    #Log operations
    wb.save(filename='ceb.xlsx')

    # Process PAY
    sheetPAY = wb['PAY']
    for row in sheetPAY.rows:
        # print('Cell {0} -- te-done {1} -- in-done{2}'.format(row[rowinx].value, 'Okay', 'Okay'))
        #
        # Robot Process 2:Process Delete Action 1 - user terminators management
        row[-1].value = 'Del action 1 done'
        # Robot Process 3:Process Delete Action 2 - user information management
        row[-2].value = 'Del action 2 done'

    # Log operations
    wb.save(filename='ceb.xlsx')

    wb.close()


    #done
    tnow = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    print ('Done {0}'.format(tnow))





