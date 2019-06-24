from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import os, sys
import xlrd
from xlutils.copy import copy
from tkinter import *
from tkinter import messagebox as mBox

root = Tk()
root.title('CEBbank robot POS manager')
Label(root, text='Username :').grid(row=0, column=0)
Label(root, text='Password :').grid(row=1, column=0)
Label(root, text='Input Path :').grid(row=2, column=0)
Label(root, text='Input file name(xls) :').grid(row=3, column=0)
Label(root, text='Output Path :').grid(row=4, column=0)
Label(root, text='Output file name(xls) :').grid(row=5, column=0)
Label(root, text='Business:').grid(row=6, column=0)

v1 = StringVar()
v2 = StringVar()
v3 = StringVar()
v4 = StringVar()
v5 = StringVar()
v6 = StringVar()
v7 = StringVar()
# pre setting
v3.set("C:/Users/Chencheng/Desktop/input/")
v4.set("input.xls")
v5.set("C:/Users/Chencheng/Desktop/output")
v6.set("output.xls")
v7.set("POS")
# define GUI
name = Entry(root, textvariable=v1)
pws = Entry(root, textvariable=v2, show='$')
inputP = Entry(root, textvariable=v3)
inputF = Entry(root, textvariable=v4)
outputP = Entry(root, textvariable=v5)
outputF = Entry(root, textvariable=v6)
bussP = Entry(root, textvariable=v7)

# GUI
name.grid(row=0, column=1, padx=10, pady=5)
pws.grid(row=1, column=1, padx=10, pady=5)
inputP.grid(row=2, column=1, padx=10, pady=5)
inputF.grid(row=3, column=1, padx=10, pady=5)
outputP.grid(row=4, column=1, padx=10, pady=5)
outputF.grid(row=5, column=1, padx=10, pady=5)
bussP.grid(row=6, column=1, padx=10, pady=5)


def _read(username, password, inputpath, inputfile, outputpath, outputfile, buss):
    # get path
    Ipath = _repalce(inputpath)
    Ifile = inputfile
    Opath = _repalce(outputpath)
    Ofile = outputfile
    # read excel
    data1 = xlrd.open_workbook(Ipath + '/' + Ifile)
    sheet1 = data1.sheet_by_name('Sheet1')
    data2 = copy(data1)
    # load first sheet
    sheet2 = data2.get_sheet(0)
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    print(nrows, " Lines")
    # background option
    # fireFoxOptions = webdriver.FirefoxOptions()
    # fireFoxOptions.set_headless()
    # driver=webdriver.Firefox(firefox_options=fireFoxOptions)
    # open browser
    driver = webdriver.Chrome()
    time.sleep(1)
    url = 'https://iam.cebbank.com/'
    driver.get(url)
    time.sleep(1)
    # login
    driver.find_element_by_xpath("//div[2]/div[1]/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/input").send_keys(
        username)
    driver.find_element_by_xpath("//div[2]/div[1]/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[3]/input").send_keys(
        password)
    driver.find_element_by_xpath("//div[2]/div[1]/div/div/div/div[2]/div/div[2]/div[2]/div[2]/div[5]").click()
    # initialize
    i = 0
    # Loop by sheet rows
    for i in range(nrows):
        print(i)
        # Sheet B1->cell(0,1)
        cell_A1 = sheet1.cell(i, 1).value
        companyid = cell_A1.strip()
        # print(companyid)
        url = 'https://www.tianyancha.com/search?key=' + companyid
        # print(url)
        driver.get(url)
        time.sleep(2)
        handles = driver.window_handles
        driver.switch_to.window(handles[0])
        try:
            # Company name
            driver.find_element_by_xpath("//div[2]/div/div[1]/div/div[3]/div/div[2]/div[1]/a").is_displayed()
        except:
            print("no companyid!")
            sheet2.write(i, 2, 'no companyid!')
        else:
            # print ("found companyid!")
            status = driver.find_element_by_xpath("//div[2]/div/div[1]/div/div[3]/div/div[2]/div[1]/div").text
            driver.find_element_by_xpath("//div[2]/div/div[1]/div/div[3]/div/div[2]/div[1]/a").click()
            time.sleep(1)
            driver.close()
            time.sleep(1)
            # print (handles)
            handles = driver.window_handles
            driver.switch_to.window(handles[0])
            # driver.implicitly_wait(5)
            time.sleep(3)
            try:
                # owner name position
                driver.find_element_by_xpath(
                    "//div[2]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/table[1]/tbody/tr[1]/td[1]/div/div[1]/div[2]/div[1]/a").is_displayed()
            except:
                # hospital
                name = driver.find_element_by_xpath(
                    "//div[2]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/table[1]/tbody/tr/td[1]").text
                type = driver.find_element_by_xpath("//div[2]/div/div[1]/div[2]/div[2]/div[4]/span").text
            else:
                name = driver.find_element_by_xpath(
                    "//div[2]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/table[1]/tbody/tr[1]/td[1]/div/div[1]/div[2]/div[1]/a").text
                type = driver.find_element_by_xpath(
                    "//div[2]/div/div[2]/div[1]/div[2]/div[2]/div[1]/div[2]/div[2]/table[2]/tbody/tr[2]/td[4]").text
            print(name, status, type)
            sheet2.write(i, 2, name)
            sheet2.write(i, 3, status)
            sheet2.write(i, 4, type)
            # save
            data2.save(Opath + '/' + Ofile)
        # driver.switch_to.window(handles[0])
    driver.quit()
    # final save
    data2.save(Opath + '/' + Ofile)
    mBox.showinfo('Python Message', 'Task is done, please have a look.')


def _quit():
    root.quit()
    root.destroy()
    exit()


def _repalce(old_str):
    new_str = old_str.replace('/', '\/')
    return new_str


# read GUI
def _check():
    if name.get() == '' or pws.get() == '':
        mBox.showinfo('Python Message', 'Please input username/password for https://iam.cebbank.com/.')
    else:
        if inputP.get() == '' or inputF.get() == '' or outputP.get() == '' or outputF.get() == '' or bussP.get() == '':
            mBox.showinfo('Python Message', 'Please input saving path.')
        else:
            username = name.get()
            password = pws.get()
            inputpath = inputP.get()
            inputfile = inputF.get()
            outputpath = outputP.get()
            outputfile = outputF.get()
            buss = bussP.get()
            _read(username, password, inputpath, inputfile, outputpath, outputfile, buss)


Button(root, text='Start', width=10, command=_check).grid(row=8, column=0, sticky=W, padx=10, pady=5)
Button(root, text='Exit', width=10, command=_quit).grid(row=8, column=4, sticky=E, padx=10, pady=5)

mainloop()

