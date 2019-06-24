import selenium as se
from selenium import webdriver

import time

def GetLoginCmdStrs(filename):
    loginCmds = GetCmdStrs("Login.prs")
    URL = loginCmds[0].split("->")[1]
    uname = loginCmds[1].split("->")[1]
    password = loginCmds[2].split("->")[1]

    strs = [URL,uname,password]
    return strs

def GetCmdStrs(filename):
    cmdStrs = []

    f = open(filename,"r")
    for row in f:
        cmdStrs.append(row)

    f.close()
    return cmdStrs

def GetXpAct(cmdStr):
    return (cmdStr.split("->"))

if __name__ == "__main__":
    #open the URL
    loginCmds = GetLoginCmdStrs("Login.prs")
    brows = webdriver.Chrome()
    brows.get(loginCmds[0])

    #Action by the Process.prs
    cmdStrs = GetCmdStrs("Process.prs")
    for cmdStr in cmdStrs:
        prsStrs = GetXpAct(cmdStr)
        print(brows.find_elements_by_xpath(prsStrs[0]).text())