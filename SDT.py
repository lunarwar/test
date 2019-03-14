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
import random
import socket
import getpass

user_name = getpass.getuser()
def weboperation(start,end,driver):
    driver.get("https://app.teleows.com")
    driver.find_element_by_id("usernameInput").send_keys("rgs336")
    driver.find_element_by_id("password").send_keys("nathan@1993")
    driver.find_element_by_id("btn_submit").click()
    driver.get(
        "https://1025-app.teleows.com/app/spl/AdvancedQuery/AdvancedQuery.spl")
    temp = driver.find_element_by_id("template")
    temp.send_keys("sdt_temp")
    sl.sleep(5)
    driver.find_element_by_class_name("sdm_list_line").click()
    sl.sleep(5)
    
    er = driver.find_element_by_id("createstarttime_input")
    for t in range(30):
        er.send_keys(Keys.BACKSPACE)
 
    # driver.find_element_by_id("createstarttime_input").send_keys("2019-01-01 00:00:00")
    
    extendq = driver.find_element_by_id("extendcondition_textarea")
    extendq.send_keys('"Last Update Time" >= "' + start + '" or "CreateTime(Create SDT)" >= "'+ start + '"')

    # extendq = driver.find_element_by_id("extendcondition_textarea")
    # extendq.send_keys('"Last Update Time" >= "' + start + '" and "Last Update Time" <= "' + end +'"')
    
    # endp = driver.find_element_by_id("createendtime_input")
    # sl.sleep(5)
    # endp.send_keys(end)  
    sl.sleep(5)        
    driver.find_element_by_id("exportqueryresult").click()
    sl.sleep(5)
    driver.switch_to.window(driver.window_handles[1])
    sl.sleep(5)

    expID = driver.find_element_by_name("exportinfoid").get_attribute('value')

    driver.find_element_by_id("ServiceButton1").click()
    # sleep and refresh
    sl.sleep(30)
    driver.find_element_by_id("TextInput1").send_keys(expID)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    sl.sleep(30)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    sl.sleep(15)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    # download file
    js = 'document.getElementsByName("nf2")[0].click();'
    driver.execute_script(js)
    print("Service desk TT file downloaded SUCCESSFULLY")
    sl.sleep(20)
    js1 = 'document.getElementsByName("nf3")[0].click();'
    driver.execute_script(js1)
    sl.sleep(10)
    print("Preparing to execute database function")
    sl.sleep(2)
    driver.close()
    sl.sleep(5)



def open_xlsx():
    _date = str(datetime.datetime.today())[:10].replace("-", "")
    path = "C:/Users/" +user_name +"/Downloads/servdesk_temp/"
    arr = os.listdir(path)
    # return a list of all files in the starting with INCIDENT TICKET 3.0_
    
    for xlsx in arr:
        if xlsx[:20] == "Service Desk Ticket_":
            insrt_file = path + xlsx
            sheetname = 'Service Desk Ticket'
            return [sheetname,insrt_file]

def renameDownloadedSDT(filz):
    _date = str(datetime.datetime.today() - timedelta(1))[:10]
    path = "C:/Users/" + user_name + "/Downloads/servdesk_temp/"
    os.rename(filz, path + _date + filz[43:])


def dbconnect():
    con = cx_Oracle.connect('datacenter/datacenter@host/orcl',encoding='UTF-8')
    return con

def data_preprocess():
    xlsxfile = open_xlsx()[1]
    df = pd.read_excel(open(xlsxfile,'rb'), sheet_name=open_xlsx()[0])
    #log = open("C:\\Users\\uset\\Desktop\\machine learning\\sdt_d.txt","a") 
    con = dbconnect()
    corsor = con.cursor()
    nas = df.columns.values  # excel sheet headers
    # viz = colunms in excel sheet
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
        nas[11],
        nas[12],
        nas[13],
        nas[14],
        nas[15],
        nas[16],
        nas[17],
        nas[18],
        nas[19],
        nas[20],
        nas[21],
        nas[22],
        nas[23],
        nas[24],
        nas[25],
        nas[26],
        nas[27],
        nas[28],
        nas[29],
        nas[30],
        nas[34],
        nas[35],
        nas[36],
        nas[37],
        nas[38],
        nas[39],
        nas[40],
        nas[41],
        nas[42],
        nas[43],
        nas[44],
        nas[45],
        nas[46],
        nas[47],
        nas[48],
        nas[49],
        nas[50],
        nas[51],
        nas[52],
        nas[53],
        nas[54],
        nas[55],
        nas[56],
        nas[57],
        nas[58],
        nas[59],
        nas[60],
        nas[61],
        nas[62],
        nas[63],
        nas[64],
        nas[65],
        nas[66],
        nas[67],
        nas[68],
        nas[69],
        nas[70],
        nas[71],
        nas[72],
        nas[73],
        nas[74],
        nas[76],
        nas[77],
        nas[78],
        nas[79],
        nas[80],
        nas[81],
        nas[82],
        nas[83],
        nas[84],
        nas[85],
        nas[86],
        nas[87],
        nas[88],
        nas[89],
        nas[90],
        nas[91],
        nas[92],
        nas[93],
        nas[94],
        nas[95],
        nas[96],
        nas[97],
        nas[98],
        nas[99],
        nas[100],
        nas[101],
        nas[102],
        nas[103],
        nas[104],
        nas[105],
        nas[106],
        nas[118],
        nas[119],
        nas[120],
        nas[121],
        nas[122],
        nas[123],
        nas[124],
        nas[125],
        nas[126],
        nas[31],
        nas[32],
        nas[33],
        nas[75],
        nas[107],
        nas[108],
        nas[109],
        nas[110],
        nas[111],
        nas[112],
        nas[113],
        nas[114],
        nas[115],
        nas[116],
        nas[117],
        nas[127],
        nas[128]

    ]]
    query = """INSERT INTO MTN_SDT_HOURLY (
                ORDERID,
                WORKFLOWTYPE,
                TOTALSLAPAUSEDURATION,
                TITLECREATESDTD,
                TICKETSTATUS,
                SLASTATUS,
                SLAPAUSEREASON,
                SLAPAUSED,
                SLA_SUSPEND_STATE,
                PARENTORDERID,
                ORIGINATOR,
                LASTUPDATETIME,
                EXEMPTED,
                ESCALATED,
                ESCALATELEVEL,
                CURRENTOPERATOR,
                CURRENTPHASE,
                CREATETIME,
                CLOSETIME,
                BUSINESSSTATUS,
                ASSOCIATEORDERID,
                ALLSUBTICKETSCLOSED,
                ABORTREASON,
                SUBMITTIMEHANDLESDT,
                OPERATORHANDLESDT,
                CREATETIMEHANDLESDT,
                OPERATIONMODEHANDLESDT,
                COPYTOHANDLESDT,
                DESCRIPTIONHANDLEDESCRIPTIONRH,
                DESCRIPTIONHANDLEDESCRIPTIONHA,
                ASSIGNTOHANDLESDT,
                SUBMITTIMECREATESDT,
                OPERATORCREATESDT,
                CREATETIMECREATESDT,
                TITLECREATESDT,
                TEMPLATECREATESDT,
                SITEISDOWNCREATESDT,
                SITEIDCREATESDT,
                SEVERITYDEFINITIONCREATESDT,
                SEVERITYCREATESDT,
                SERVICEIMPACTINGCREATESDT,
                CREATEREQUESTTYPECREATESDT,
                ASSOCIATETTIDCREATESDT,
                ORDERIDCREATESDT,
                IVRCALLLEVELCREATESDT,
                EVENTCATEGORYCREATESDT,
                DIRECTASSOCIATETICKETSQTYCREAT,
                DIRECTASSOCIATETICKETSCREATESD,
                TICKETLEVELCREATESDT,
                SERVICETYPECREATESDT,
                REQUESTTYPECREATESDT,
                PLANNEDINCREATESDT,
                REQUESTSOURCETICKETIDCREATESDT,
                REQUESTSOURCECREATESDT,
                REQUESTORPHONECREATESDT,
                REQUESTORNAMECREATESDT,
                REQUESTORMOBILEPHONECREATESDT,
                REQUESTOREMAILCREATESDT,
                PLANNEDOUTCREATESDT,
                REQUESTDETAILSCREATESDT,
                REGIONCREATESDT,
                FAULTSERVICENOCREATESDT,
                FAULTOCCURTIMECREATESDT,
                COPYTOCREATESDT,
                ATTACHMENTCREATESDT,
                ASSIGNTOCREATESDT,
                ADDRESSCREATESDT,
                BSCRNCCREATESDT,
                ALLASSOCIATETICKETSCREATESDT,
                ALLASSOCIATETICKETQTYCREATESDT,
                ALARMCLASSFICATIONCREATESDT,
                ALARMNAMECREATESDT,
                WOSTATUSPROCESSSDT,
                WFMWOIDPROCESSSDT,
                WFMRCAPROCESSSDT,
                ATTACHMENTUPDATEATTACHMENTPROC,
                SUBMITTIMEPROCESSSDT,
                OPERATORPROCESSSDT,
                CREATETIMEPROCESSSDT,
                RCACATEGORYPROCESSSDT,
                CREATEREQSTTYPECREATESDT,
                RCAPROCESSSDT,
                SOLUTIONTYPEPROCESSSDT,
                RESOLVENTDESCRIPTIONPROCESSSDT,
                SOLUTIONDESCRIPTIONPROCESSSDT,
                REQUESTCHILDTYPERESOLVEREQUEST,
                REQUESTCHILDTYPERESOLVEREQUTD,
                PRIORITYRESOLVEPRIORITYRPROCES,
                PRIORITYRESOLVEPRIORITYPROCESS,
                ACTUALOUTPROCESSSDT,
                ACTUALINPROCESSSDT,
                COPYTORESOLVECOPYTOPROCESSSDT,
                REASONTYPEPROCESSSDT,
                REASDESCRIPTIONRESOLVECAUSED,
                REASDESCRIPTIONRESOLVECAUSEDD,
                ATTACHMENTRESOLVEATTACHMENTPRO,
                ASSIGNTOPROCESSSDT,
                REMARKPROCESSSDT,
                OPERATIONMODEPROCESSSDT,
                FAULTRECOVERYTIMEPROCESSSDT,
                ESCALATIONDETAILSPROCESSSDT,
                REQSTTYPECREATESDT,
                ALARMNAMEPROCESSSDT,
                SUBMITTIMECONFIRMSDT,
                OPERATORCONFIRMSDT,
                CREATETIMECONFIRMSDT,
                OPERATIONMODECONFIRMSDT,
                ATTACHMENTCONFIRMSDT,
                VALIDATETIMECONFIRMSDT,
                SATISFACTORYCONFIRMSDT,
                DESCRIPTIONCLOSEDESCRIPTIONRCO,
                DESCRIPTIONCLOSEDESCRIPTIONACO,
                last_update_time,
                OLA_2_TIMEHANDLESDT,
                OLA_2_OPERATORHANDLESDT,
                OLA_2_DURATIONHANDLESDT,
                OLA_1_TIMECREATESDT,
                TTRPROCESSSDT,
                TTRSEPROCESSSDT,
                REASONFORREJECTPROCESSSDT,
                OLA_4_TIMEPROCESSSDT,
                OLA_4_OPERATORPROCESSSDT,
                OLA_4_DURATIONPROCESSSDT,
                OLA_3_TIMEPROCESSSDT,
                OLA_3_OPERATORPROCESSSDT,
                OLA_3_DURATIONPROCESSSDT,
                OLA_2_TIMEHANDLESDT2,
                FAULTOCCURTIMEPROCESSSD,
                OLA_5_TIMECONFIRMSDTD,
                OLA_5_OPERATORCONFIRMSDT
                    ) VALUES """
    insert_date =  str(datetime.datetime.today())[:19]
    for i in range(0, len(viz)):

        data = (
            str(viz[nas[0]][i]).replace("'", ""),
            str(viz[nas[1]][i]).replace("'", ""),
            str(viz[nas[2]][i]).replace("'", ""),
            str(viz[nas[3]][i]).replace("'", ""),
            str(viz[nas[4]][i]).replace("'", ""),
            str(viz[nas[5]][i]).replace("'", ""),
            str(viz[nas[6]][i]).replace("'", ""),
            str(viz[nas[7]][i]).replace("'", ""),
            str(viz[nas[8]][i]).replace("'", ""),
            str(viz[nas[9]][i]).replace("'", ""),
            str(viz[nas[10]][i]).replace("'", ""),
            str(viz[nas[11]][i]).replace("'", ""),
            str(viz[nas[12]][i]).replace("'", ""),
            str(viz[nas[13]][i]).replace("'", ""),
            str(viz[nas[14]][i]).replace("'", ""),
            str(viz[nas[15]][i]).replace("'", ""),
            str(viz[nas[16]][i]).replace("'", ""),
            str(viz[nas[17]][i]).replace("'", ""),
            str(viz[nas[18]][i]).replace("'", ""),
            str(viz[nas[19]][i]).replace("'", ""),
            str(viz[nas[20]][i]).replace("'", ""),
            str(viz[nas[21]][i]).replace("'", ""),
            str(viz[nas[22]][i]).replace("'", ""),
            str(viz[nas[23]][i]).replace("'", ""),
            str(viz[nas[24]][i]).replace("'", ""),
            str(viz[nas[25]][i]).replace("'", ""),
            str(viz[nas[26]][i]).replace("'", ""),
            str(viz[nas[27]][i]).replace("'", ""),
            str(viz[nas[28]][i]).replace("'", ""),
            str(viz[nas[29]][i]).replace("'", ""),
            str(viz[nas[30]][i]).replace("'", ""),
            str(viz[nas[34]][i]).replace("'", ""),
            str(viz[nas[35]][i]).replace("'", ""),
            str(viz[nas[36]][i]).replace("'", ""),
            str(viz[nas[37]][i]).replace("'", ""),
            str(viz[nas[38]][i]).replace("'", ""),
            str(viz[nas[39]][i]).replace("'", ""),
            str(viz[nas[40]][i]).replace("'", ""),
            str(viz[nas[41]][i]).replace("'", ""),
            str(viz[nas[42]][i]).replace("'", ""),
            str(viz[nas[43]][i]).replace("'", ""),
            str(viz[nas[44]][i]).replace("'", ""),
            str(viz[nas[45]][i]).replace("'", ""),
            str(viz[nas[46]][i]).replace("'", ""),
            str(viz[nas[47]][i]).replace("'", ""),
            str(viz[nas[48]][i]).replace("'", ""),
            str(viz[nas[49]][i]).replace("'", ""),
            str(viz[nas[50]][i]).replace("'", ""),
            str(viz[nas[51]][i]).replace("'", ""),
            str(viz[nas[52]][i]).replace("'", ""),
            str(viz[nas[53]][i]).replace("'", ""),
            str(viz[nas[54]][i]).replace("'", ""),
            str(viz[nas[55]][i]).replace("'", ""),
            str(viz[nas[56]][i]).replace("'", ""),
            str(viz[nas[57]][i]).replace("'", ""),
            str(viz[nas[58]][i]).replace("'", ""),
            str(viz[nas[59]][i]).replace("'", ""),
            str(viz[nas[60]][i]).replace("'", ""),
            str(viz[nas[61]][i]).replace("'", ""),
            str(viz[nas[62]][i]).replace("'", ""),
            str(viz[nas[63]][i]).replace("'", ""),
            str(viz[nas[64]][i]).replace("'", ""),
            str(viz[nas[65]][i]).replace("'", ""),
            str(viz[nas[66]][i]).replace("'", ""),
            str(viz[nas[67]][i]).replace("'", ""),
            str(viz[nas[68]][i]).replace("'", ""),
            str(viz[nas[69]][i]).replace("'", ""),
            str(viz[nas[70]][i]).replace("'", ""),
            str(viz[nas[71]][i]).replace("'", ""),
            str(viz[nas[72]][i]).replace("'", ""),
            str(viz[nas[73]][i]).replace("'", ""),
            str(viz[nas[74]][i]).replace("'", ""),
            str(viz[nas[76]][i]).replace("'", ""),
            str(viz[nas[77]][i]).replace("'", ""),
            str(viz[nas[78]][i]).replace("'", ""),
            str(viz[nas[79]][i]).replace("'", ""),
            str(viz[nas[80]][i]).replace("'", ""),
            str(viz[nas[81]][i]).replace("'", ""),
            str(viz[nas[82]][i]).replace("'", ""),
            str(viz[nas[83]][i]).replace("'", ""),
            str(viz[nas[84]][i]).replace("'", ""),
            str(viz[nas[85]][i]).replace("'", ""),
            str(viz[nas[86]][i]).replace("'", ""),
            str(viz[nas[87]][i]).replace("'", ""),
            str(viz[nas[88]][i]).replace("'", ""),
            str(viz[nas[89]][i]).replace("'", ""),
            str(viz[nas[90]][i]).replace("'", ""),
            str(viz[nas[91]][i]).replace("'", ""),
            str(viz[nas[92]][i]).replace("'", ""),
            str(viz[nas[93]][i]).replace("'", ""),
            str(viz[nas[94]][i]).replace("'", ""),
            str(viz[nas[95]][i]).replace("'", ""),
            str(viz[nas[96]][i]).replace("'", ""),
            str(viz[nas[97]][i]).replace("'", ""),
            str(viz[nas[98]][i]).replace("'", ""),
            str(viz[nas[99]][i]).replace("'", ""),
            str(viz[nas[100]][i]).replace("'", ""),
            str(viz[nas[101]][i]).replace("'", ""),
            str(viz[nas[102]][i]).replace("'", ""),
            str(viz[nas[103]][i]).replace("'", ""),
            str(viz[nas[104]][i]).replace("'", ""),
            str(viz[nas[105]][i]).replace("'", ""),
            str(viz[nas[106]][i]).replace("'", ""),
            str(viz[nas[118]][i]).replace("'", ""),
            str(viz[nas[119]][i]).replace("'", ""),
            str(viz[nas[120]][i]).replace("'", ""),
            str(viz[nas[121]][i]).replace("'", ""),
            str(viz[nas[122]][i]).replace("'", ""),
            str(viz[nas[123]][i]).replace("'", ""),
            str(viz[nas[124]][i]).replace("'", ""),
            str(viz[nas[125]][i]).replace("'", ""),
            str(viz[nas[126]][i]).replace("'", ""),
            insert_date,
            str(viz[nas[31]][i]).replace("'", ""),
            str(viz[nas[32]][i]).replace("'", ""),
            str(viz[nas[33]][i]).replace("'", ""),
            str(viz[nas[75]][i]).replace("'", ""),
            str(viz[nas[107]][i]).replace("'", ""),
            str(viz[nas[108]][i]).replace("'", ""),
            str(viz[nas[109]][i]).replace("'", ""),
            str(viz[nas[110]][i]).replace("'", ""),
            str(viz[nas[111]][i]).replace("'", ""),
            str(viz[nas[112]][i]).replace("'", ""),
            str(viz[nas[113]][i]).replace("'", ""),
            str(viz[nas[114]][i]).replace("'", ""),
            str(viz[nas[115]][i]).replace("'", ""),
            str(viz[nas[116]][i]).replace("'", ""),
            str(viz[nas[117]][i]).replace("'", ""),
            str(viz[nas[127]][i]).replace("'", ""),
            str(viz[nas[128]][i]).replace("'", "")
        )
        try: 
            corsor.execute(query + str(data).replace("nan","''").replace('"',""))
            print("inserted successfully")
        except Exception as f: 
            #print(f)  
            try:
                corsor.execute(query + str(str(data).replace("nan","''").replace('"',"").encode("utf-8")))
                print("inserted in second try statment")
            except Exception as d:
                #print(d)
                try:
                    lst = list(data)
                    lst[25] = "contact ictom team"
                    lst[26] = "contact ictom team"
                    data = tuple(lst)
                    corsor.execute(query + str(str(data).replace("nan","''").replace('"',"").encode("utf-8")))
                    print("inserted on third exception")
                except Exception as e:
                    #print("errrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrro")
                    #print(str(data).replace("nan","''").replace('"',""))  
                    log.write(str(datetime.datetime.today())[:18] + '\n')
                    log.write(str(e) + '\n')
                    log.write(str(viz["OrderId"][i]) + '\n') 
                    log.write(str(datetime.datetime.today())[:18] + '\n') 
                #continue
                performupdate(data,corsor) 
            continue
    corsor.close()
    con.commit() 
    print("task completed")
    log.close()
    renameDownloadedSDT(xlsxfile)


def performupdate(dataset,corsor):
    query = [
        "ORDERID",
        "WORKFLOWTYPE",
        "TOTALSLAPAUSEDURATION",
        "TITLECREATESDTD",
        "TICKETSTATUS",
        "SLASTATUS",
        "SLAPAUSEREASON",
        "SLAPAUSED",
        "SLA_SUSPEND_STATE",
        "PARENTORDERID",
        "ORIGINATOR",
        "LASTUPDATETIME",
        "EXEMPTED",
        "ESCALATED",
        "ESCALATELEVEL",
        "CURRENTOPERATOR",
        "CURRENTPHASE",
        "CREATETIME",
        "CLOSETIME",
        "BUSINESSSTATUS",
        "ASSOCIATEORDERID",
        "ALLSUBTICKETSCLOSED",
        "ABORTREASON",
        "SUBMITTIMEHANDLESDT",
        "OPERATORHANDLESDT",
        "CREATETIMEHANDLESDT",
        "OPERATIONMODEHANDLESDT",
        "COPYTOHANDLESDT",
        "DESCRIPTIONHANDLEDESCRIPTIONRH",
        "DESCRIPTIONHANDLEDESCRIPTIONHA",
        "ASSIGNTOHANDLESDT",
        "SUBMITTIMECREATESDT",
        "OPERATORCREATESDT",
        "CREATETIMECREATESDT",
        "TITLECREATESDT",
        "TEMPLATECREATESDT",
        "SITEISDOWNCREATESDT",
        "SITEIDCREATESDT",
        "SEVERITYDEFINITIONCREATESDT",
        "SEVERITYCREATESDT",
        "SERVICEIMPACTINGCREATESDT",
        "CREATEREQUESTTYPECREATESDT",
        "ASSOCIATETTIDCREATESDT",
        "ORDERIDCREATESDT",
        "IVRCALLLEVELCREATESDT",
        "EVENTCATEGORYCREATESDT",
        "DIRECTASSOCIATETICKETSQTYCREAT",
        "DIRECTASSOCIATETICKETSCREATESD",
        "TICKETLEVELCREATESDT",
        "SERVICETYPECREATESDT",
        "REQUESTTYPECREATESDT",
        "PLANNEDINCREATESDT",
        "REQUESTSOURCETICKETIDCREATESDT",
        "REQUESTSOURCECREATESDT",
        "REQUESTORPHONECREATESDT",
        "REQUESTORNAMECREATESDT",
        "REQUESTORMOBILEPHONECREATESDT",
        "REQUESTOREMAILCREATESDT",
        "PLANNEDOUTCREATESDT",
        "REQUESTDETAILSCREATESDT",
        "REGIONCREATESDT",
        "FAULTSERVICENOCREATESDT",
        "FAULTOCCURTIMECREATESDT",
        "COPYTOCREATESDT",
        "ATTACHMENTCREATESDT",
        "ASSIGNTOCREATESDT",
        "ADDRESSCREATESDT",
        "BSCRNCCREATESDT",
        "ALLASSOCIATETICKETSCREATESDT",
        "ALLASSOCIATETICKETQTYCREATESDT",
        "ALARMCLASSFICATIONCREATESDT",
        "ALARMNAMECREATESDT",
        "WOSTATUSPROCESSSDT",
        "WFMWOIDPROCESSSDT",
        "WFMRCAPROCESSSDT",
        "ATTACHMENTUPDATEATTACHMENTPROC",
        "SUBMITTIMEPROCESSSDT",
        "OPERATORPROCESSSDT",
        "CREATETIMEPROCESSSDT",
        "RCACATEGORYPROCESSSDT",
        "CREATEREQSTTYPECREATESDT",
        "RCAPROCESSSDT",
        "SOLUTIONTYPEPROCESSSDT",
        "RESOLVENTDESCRIPTIONPROCESSSDT",
        "SOLUTIONDESCRIPTIONPROCESSSDT",
        "REQUESTCHILDTYPERESOLVEREQUEST",
        "REQUESTCHILDTYPERESOLVEREQUTD",
        "PRIORITYRESOLVEPRIORITYRPROCES",
        "PRIORITYRESOLVEPRIORITYPROCESS",
        "ACTUALOUTPROCESSSDT",
        "ACTUALINPROCESSSDT",
        "COPYTORESOLVECOPYTOPROCESSSDT",
        "REASONTYPEPROCESSSDT",
        "REASDESCRIPTIONRESOLVECAUSED",
        "REASDESCRIPTIONRESOLVECAUSEDD",
        "ATTACHMENTRESOLVEATTACHMENTPRO",
        "ASSIGNTOPROCESSSDT",
        "REMARKPROCESSSDT",
        "OPERATIONMODEPROCESSSDT",
        "FAULTRECOVERYTIMEPROCESSSDT",
        "ESCALATIONDETAILSPROCESSSDT",
        "REQSTTYPECREATESDT",
        "ALARMNAMEPROCESSSDT",
        "SUBMITTIMECONFIRMSDT",
        "OPERATORCONFIRMSDT",
        "CREATETIMECONFIRMSDT",
        "OPERATIONMODECONFIRMSDT",
        "ATTACHMENTCONFIRMSDT",
        "VALIDATETIMECONFIRMSDT",
        "SATISFACTORYCONFIRMSDT",
        "DESCRIPTIONCLOSEDESCRIPTIONRCO",
        "DESCRIPTIONCLOSEDESCRIPTIONACO",
        "last_update_time",
        "OLA_2_TIMEHANDLESDT",
        "OLA_2_OPERATORHANDLESDT",
        "OLA_2_DURATIONHANDLESDT",
        "OLA_1_TIMECREATESDT",
        "TTRPROCESSSDT",
        "TTRSEPROCESSSDT",
        "REASONFORREJECTPROCESSSDT",
        "OLA_4_TIMEPROCESSSDT",
        "OLA_4_OPERATORPROCESSSDT",
        "OLA_4_DURATIONPROCESSSDT",
        "OLA_3_TIMEPROCESSSDT",
        "OLA_3_OPERATORPROCESSSDT",
        "OLA_3_DURATIONPROCESSSDT",
        "OLA_2_TIMEHANDLESDT2",
        "FAULTOCCURTIMEPROCESSSD",
        "OLA_5_TIMECONFIRMSDTD",
        "OLA_5_OPERATORCONFIRMSDT"
    ]
    #con = dbconnect()
    #corsor = con.cursor()
    for i,j in enumerate(query):
        strg = "UPDATE MTN_SDT_HOURLY SET " + str(j) + " = '" + str(dataset[i]).replace("nan","''").replace('"',"") + "'" + " where ORDERID = '" + str(dataset[0]) + "'"
        try:
            corsor.execute(strg)
        except:
            continue
    print("updated")
def main(date1,date2):

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : 'C:/Users/' +user_name+ '/Downloads/servdesk_temp'}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    weboperation(date1,date2,driver)
    return

if __name__ =='__main__':

    logp = open("C:\\Users\\" + user_name +"\\Desktop\\machine learning\\owsobotlog.txt", "a")
    password = getpass.getpass(prompt='Enter Password: ', stream=None)
    print(user_name)
    if password == str(user_name):
        username = socket.getfqdn()
        start_time = str(datetime.datetime.now() - datetime.timedelta(days=2))[:19]
        end_time =   str(datetime.datetime.today())[:10]  + " 23:59:59"
        logp.write("OWSOBOT APP STARTED SUCCESSFULLY" + '\n')
        logp.write("HOST NAME: " + str(username) + '\n')
        logp.write("Execution Date: " + str(datetime.datetime.now()) + '\n')
        try:
            main(start_time, end_time)
            logp.write("DATA EXTRACTION STATUS: Successful"  + '\n')
            try:
                data_preprocess()
                logp.write("DATA INSERT/UPDATE ORACLE: Successful" + '\n')
            except Exception as EX:
                logp.write("DATA EXTRACTION: "+ str(EX) + '\n')
                logp.write("DATA INSERT/UPDATE ORACLE: failed" + '\n')
        except Exception as e:
            logp.write("DATA EXTRACTION: "+ str(e) + '\n')
            logp.write("DATA EXTRACTION: Failed" + '\n')
            try:
                data_preprocess()
                logp.write("DATA INSERT/UPDATE ORACLE: Successful" + '\n')
            except Exception as EX:
                logp.write("DATA EXTRACTION: "+ str(EX) + '\n')
                logp.write("DATA INSERT/UPDATE ORACLE: failed" + '\n')
    else:
        logp.write("WRONG PASSWORD " + '\n')

