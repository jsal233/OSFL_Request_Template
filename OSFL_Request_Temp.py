"""
Coded by: Jayson Salas
Date: 03/14/2022
Title: OSFL Request Template generator
"""

import time, os, re
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# use for doubleclick action
# from selenium.webdriver import ActionChains
from openpyxl import load_workbook, Workbook

def logit(x):
    """Use for logging"""


    log_fname = ("log_OSFL_RT_" + date + ".txt")
    log_fpath = os.path.join(log_dir, log_fname)

    f = open(log_fpath, "a")
    f.write(ftime + " :: " + x + "\n")
    f.close()

def mytime():
    # Get the current date and time.
    time_lt = time.localtime()  # get struct_time
    ftime = time.strftime("%m-%d-%Y %H:%M:%S", time_lt)
    date = time.strftime("%m-%d-%Y", time_lt)

    return date, ftime

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # Set Dates
    date, ftime = mytime()

    # Set the directories
    curr_dir = os.getcwd()
    print(curr_dir)

    # log directory
    log_dir = (curr_dir + "/logs/")
    # create a directory if not exists
    try:
        os.mkdir(log_dir)
    except FileExistsError:
        print("Log directory already exists")

    # output directory
    out_dir = (curr_dir + "/output/")
    # create a directory if not exists
    try:
        os.mkdir(out_dir)
    except FileExistsError:
        print("output directory already exists")

    logit("-" * 40 + " Start execution " + "-" * 40)

    # Create output file for writing
    # Create xl workbook
    ORTxl_fp = (out_dir + 'OSFL_Request_Template_' + date + '.xlsx')

    # To open the workbook object is created
    ORTwb_obj = Workbook()

    # Get workbook active sheet object from the active attribute
    ORTws = ORTwb_obj.active
    ORTws.append(["Store #", "City Name", "Division", "Banner", "Division ID", "ROG", "ddssss", "Value", "MQ"])

    # Use Selenium for web mine
    ChromePATH = curr_dir + '\driver\chromedriver.exe'
    site = ("http://operations.safeway.com/sinfo/index.cgi?store=2176&misc=")
    logit(site)
    driver = webdriver.Chrome(ChromePATH)
    # action = ActionChains(driver)
    driver.get(site)

    # Iterate input file
    stores = open("store_input.txt", "r")

    for store in stores:
        store = store.rstrip()
        store = store.zfill(4)

        # Enter Store Name
        sleep(3)
        search = driver.find_element_by_xpath('//*[@id="store"]')
        # action.double_click(search).perform()
        sleep(2)
        search.send_keys(store)
        sleep(1)
        search.send_keys(Keys.ENTER)
        logit("%s store Entered" % store)

        address = driver.find_element_by_xpath('/html/body/div[3]/div/table[2]/tbody/tr[2]/td[1]')
        # c_content = "".join([c.text for c in city])
        # get only the city from the address line 2
        add = (address.text).split("\n")
        x = re.search("^([A-Z|a-z]+(\s[A-Z|a-z]+)*)\s([A-Z]{2}\s\d{5})$", add[1])
        city = x.group(1)

        sleep(1)
        division = driver.find_element_by_xpath('/html/body/div[3]/div/table[1]/tbody/tr[2]/td[6]')
        division = (division.text)
        # print("Division: %s" % division)

        sleep(1)
        banner = driver.find_element_by_xpath('/html/body/div[3]/div/table[1]/tbody/tr[2]/td[2]')
        banner = (banner.text)
        # print("Banner: %s" % banner)

        sleep(1)
        divid = driver.find_element_by_xpath('/html/body/div[3]/div/table[1]/tbody/tr[2]/td[4]')
        divid = (divid.text).zfill(2)
        # print("Division ID: %s" % divid)

        sleep(1)
        rog = driver.find_element_by_xpath('/html/body/div[3]/div/table[1]/tbody/tr[2]/td[7]')
        rog = (rog.text)
        # print("ROG: %s" % rog)

        ddssss = divid + store
        # print("ddssss: %s" % ddssss)

        valuebd = banner + "-" + divid
        # print("Value: %s" % valuebd)

        mq = "RHL_MQ_OSFL X" + ddssss
        # print("Value: %s" % mq)

        # write the values mined
        ORTws.append([store, city, division, banner, divid, rog, ddssss, valuebd, mq])

    # save the excel file
    ORTwb_obj.save(ORTxl_fp)

    logit("-" * 40 + " End execution " + "-" * 40)
    driver.close()



