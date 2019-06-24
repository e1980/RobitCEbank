import selenium as se
from selenium import webdriver

import time
import json


def GetCmdDict(filename):
    cmdStrs = {}

    f = open(filename,"r")
    for row in f:
        row = row.split("->")
        ky = row[0]
        val = row[1]
        cmdStrs[ky]=val

    f.close()
    return cmdStrs

if __name__ == "__main__":
    #open the URL
    loginCmds = GetCmdDict("Login.prs")
    robotCmds = GetCmdDict("Process.prs")
    brows = webdriver.Chrome()
    brows.get(loginCmds['URL'])
    time.sleep(2)

    #Action by the Process.prs
    #Robot Process 1:Login 163 mail
    #robot_switch (click); robot_emailid (sendkeys); robot_password(sendkeys);robot_login(click)
    brows.find_element_by_xpath(xpath=robotCmds['robot_switch']).click()
    #brows.find_element_by_xpath(xpath=robotCmds['robot_passwd']).send_keys(loginCmds['password'])

