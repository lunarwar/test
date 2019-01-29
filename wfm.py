import pandas as pd 
import matplotlib.pyplot as plt
import datetime
import cx_Oracle
import re
import time as sl
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
import datetime
from datetime import date, timedelta
import os
import threading

def weboperation(start,end,driver):
    driver.get("https://app.teleows.com")
    driver.find_element_by_id("usernameInput").send_keys("****")
    driver.find_element_by_id("password").send_keys("***@****")
    driver.find_element_by_id("btn_submit").click()
    driver.get(
        "view-source:https://101b-app.teleows.com/app/101b/spl/WFM_AdvancedQuery/WFM_AdvancedQuery.spl")
    temp = driver.find_element_by_id("template")
    temp.send_keys("wfm_query")
    sl.sleep(5)
    driver.find_element_by_class_name("sdm_list_line").click()
    extendq = driver.find_element_by_id("extendcondition_textarea")
    extendq.send_keys('"Last Update Time" >= "' + start + '" or "Create Time" >= "'+ start + '"')

    sl.sleep(5)        
    driver.find_element_by_id("checkExport").click()
    sl.sleep(5)
    driver.switch_to.window(driver.window_handles[1])
    sl.sleep(5)
    driver.close()


def open_xlsx():
    _date = str(datetime.datetime.today())[:10].replace("-", "")
    path = "C:/Users/ewx510986/Downloads/MS_WFM/"
    arr = os.listdir(path)
    # return a list of all files in the starting with INCIDENT TICKET 3.0_
    for xlsx in arr:
        if xlsx[:16] == "Incident Ticket_":
            insrt_file = path + xlsx
            sheetname = 'Incident Ticket'
            return [sheetname,insrt_file]

def renameDownloadedSDT(filz):
    _date = str(datetime.datetime.today() - timedelta(1))[:10]
    path = "C:/Users/ewx510986/Downloads/MS_WFM/"
    os.rename(filz, path + _date + filz[46:])

def dbconnect():
    con = cx_Oracle.connect('dbusername/dbpassword@dbserveraddr/dbservice_name')
    return con

def data_preprocess():
    xlsxfile = open_xlsx()[1]
    df = pd.read_excel(open(xlsxfile,'rb'), sheet_name=open_xlsx()[0])
    log = open('airtel_itt_log.txt','a') 
    con = dbconnect()
    corsor = con.cursor()
    nas = df.columns.values  # excel sheet headers
    viz = df[[  
        nas[0],
        nas[1],
        nas[2],
        nas[3],
        nas[4],
        nas[5],
        nas[6],
        nas[7],
        nas[8],
        nas[9],
        nas[10],
        nas[11]
            ]]
    query = """INSERT INTO MS_WFM(
                TASK_ID,
                OPERATE_TYPE,
                TASK_TYPE,
                TITLE,
                TASK_STATUS,
                FM_OFFICE,
                ASSIGN_TO_FME,
                SITE_ID,
                PROJECT,
                CREATE_TIME,
                CONFIRM_TIME.
                SLA_STATUS,
                LAST_UPDATE_TIME
                    ) VALUES """
    for i in range(0, len(viz)):
        insert_date =  str(datetime.datetime.today())[:19]
        data = (
        viz[nas[0]][i],
        viz[nas[1]][i],
        str(viz[nas[2]][i]).replace("'",""),
        viz[nas[3]][i],
        viz[nas[4]][i],
        viz[nas[5]][i],
        viz[nas[6]][i],
        viz[nas[7]][i],
        viz[nas[8]][i],
        viz[nas[9]][i],
        viz[nas[10]][i],
        viz[nas[11]][i],
        insert_date 
            )

        try: 
            corsor.execute(query + str(data).replace("nan","''").replace('"',"").replace("\\"," ").replace("  "," "))
            print("insert successfully")
        except Exception as e:
            log.write(str(e) + '\n')
            log.write(str(e) + '\n')
            log.write(str(viz["OrderId"][i]) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n') 
            Performace_update(data,corsor)
            continue

    corsor.close()
    con.commit() 
    print("task completed")
    log.write(str(datetime.datetime.today())[:18] + '\n')
    log.close()
    renameDownloadedSDT(xlsxfile)

def Performace_update(dataset,corsor):
    query = [
        "TASK_ID",
        "OPERATE_TYPE",
        "TASK_TYPE",
        "TITLE",
        "TASK_STATUS",
        "FM_OFFICE",
        "ASSIGN_TO_FME",
        "SITE_ID",
        "PROJECT",
        "CREATE_TIME",
        "CONFIRM_TIME",
        "SLA_STATUS",
        "LAST_UPDATE_TIME"
    ]
    for i,j in enumerate(query):
        strg = "UPDATE AIRTEL_REPORTING_TT SET " + str(j) + " = '" + str(dataset[i]).replace("nan","''").replace('"',"").replace("\\"," ").replace("  "," ") + "'" + " where ORDERID = '" + str(dataset[0]) + "'"
        try:
            corsor.execute(strg)
        except:
            continue
    print("updated")


def main(date1,date2):
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : 'C:/Users/ewx510986/Downloads/MS_WFM'}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    weboperation(date1,date2,driver)
    return


if __name__ =='__main__':
    log = open("C:\\Users\\ewx510986\\Desktop\\machine learning\\MS_WFM.txt","w")
    _date = str(datetime.datetime.now() - datetime.timedelta(minutes=60))[:19]
    start = _date
    end  =  _date
    log.write(str(datetime.datetime.now()) + '\n') 
    main(start,end)
    data_preprocess()

    
