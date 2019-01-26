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
    driver.get("https://url")
    driver.find_element_by_id("usernameInput").send_keys("test@1")
    driver.find_element_by_id("password").send_keys("test@1")
    driver.find_element_by_id("btn_submit").click()
    driver.get(
        "url")
    temp = driver.find_element_by_id("template")
    temp.send_keys("airtel_tt_dump")
    sl.sleep(5)
    driver.find_element_by_class_name("sdm_list_line").click()
    sl.sleep(5)
    
    er = driver.find_element_by_id("createstarttime_input")
    for t in range(30):
        er.send_keys(Keys.BACKSPACE)
 
    driver.find_element_by_id("createstarttime_input").send_keys("2019-01-01 00:00:00")

    extendq = driver.find_element_by_id("extendcondition_textarea")
    extendq.send_keys('"Last Update Time" >= "' + start + '" or "CreateTime(Create TT)" >= "'+ start + '"')

    sl.sleep(5)        
    driver.find_element_by_id("exportqueryresult").click()
    sl.sleep(5)
    driver.switch_to.window(driver.window_handles[1])
    sl.sleep(5)

    expID = driver.find_element_by_name("exportinfoid").get_attribute('value')

    driver.find_element_by_id("ServiceButton1").click()
    # sleep and refresh
    sl.sleep(120)
    driver.find_element_by_id("TextInput1").send_keys(expID)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    sl.sleep(60)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    sl.sleep(45)
    driver.find_element_by_id("toolbarSearchButton").click()
    print("first refresh..........second in a minute")
    # download file
    js = 'document.getElementsByName("nf2")[0].click();'
    driver.execute_script(js)
    print("INCIDENT TT file downloaded SUCCESSFULLY")
    sl.sleep(30)
    js1 = 'document.getElementsByName("nf3")[0].click();'
    driver.execute_script(js1)
    sl.sleep(30)
    print("Preparing to execute database function")
    sl.sleep(2)
    #driver.close()


def open_xlsx():
    _date = str(datetime.datetime.today())[:10].replace("-", "")
    path = "C:/Users/ewx510986/Downloads/AIRTEL_REPORT_TT/"
    arr = os.listdir(path)
    # return a list of all files in the starting with INCIDENT TICKET 3.0_
    for xlsx in arr:
        if xlsx[:16] == "Incident Ticket_":
            insrt_file = path + xlsx
            sheetname = 'Incident Ticket'
            return [sheetname,insrt_file]

def renameDownloadedSDT(filz):
    _date = str(datetime.datetime.today() - timedelta(1))[:10]
    path = "C:/Users/ewx510986/Downloads/AIRTEL_REPORT_TT/"
    os.rename(filz, path + _date + filz[46:])

def dbconnect():
    con = cx_Oracle.connect('dbusername/dbpassword@dbserveraddr/dbservice_name')
    return con

# def clear_records(date1,date2):
#     con = dbconnect()
#     corsor = con.cursor()
#     query = "delete from AIRTEL_REPORTING_TT  WHERE TO_DATE(CREATETIME, 'YYYY/MM/DD HH24:MI:SS') <= TO_DATE('"+ date2 + " 00:00:00', 'YYYY/MM/DD HH24:MI:SS') and TO_DATE(CREATETIME, 'YYYY/MM/DD HH24:MI:SS') >= TO_DATE('"+ date1 + " 00:00:00', 'YYYY/MM/DD HH24:MI:SS')"
#     corsor.execute(query)
#     print(query)
#     corsor.close()
#     con.commit() 
#     print("initial records cleaned")
#     sl.sleep(20)

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
        nas[31],
        nas[32],
        nas[33],
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
        nas[75],
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
        nas[118],
        nas[119],
        nas[120],
        nas[121],
        nas[122],
        nas[123],
        nas[124],
        nas[129],
        nas[130],
        nas[131],
        nas[132],
        nas[133],
        nas[134],
        nas[135],
        nas[136],
        nas[137],
        nas[138],
        nas[139],
        nas[140],
        nas[141],
        nas[142],
        nas[125],
        nas[126],
        nas[127],
        nas[128]
            ]]
    query = """INSERT INTO AIRTEL_REPORTING_TT(
                OrderId,
                WorkflowType,
                SummaryCreateTT,
                TicketStatus,
                BusinessStatus,
                SLAStatus,
                CurrentOperator,
                OriginatorProcessTT,
                CreateTime,
                CloseTime,
                CurrentPhase,
                Escalated,
                Exempted,
                DescriptioncloseacceptdesConf,
                AttachmentConfirmTT,
                DescriptioncloserejectdesConf,
                SatisfactionDegreeConfirmTT,
                OperationConfirmTT,
                ParentOrderIDConfirmTT_,
                CreateTimeConfirmTT,
                OperatorConfirmTT,
                SubmitTimeConfirmTT,
                AffectedRegionCreateTT,
                AlarmIDCreateTT,
                AlarmNameCreateTT,
                AlarmTypeCreateTT,
                BSCCreateTT,
                BSC_RNCCreateTT,
                AssignToCreateTT,
                AchmentCreateTT,
                CopyToCreateTT,
                DescriptionCreateTT,
                SolutionCreateTT,
                DeviceCreateTT,
                DomainCreateTT,
                EMSCreateTT,
                EMSAlarmNoCreateTT,
                EventCodeCreateTT,
                AcknowledgementTimeCreateTT,
                CausedByCRCreateTT,
                FaultFirstOccurTimeCreateTT,
                FaultLastOccurTimeCreateTT,
                FaultLevelCreateTT,
                FaultNoCreateTT,
                TargetTimeCreateTT,
                ImpactCreateTT,
                ImpactedServiceCreateTT,
                OutageSiteQTYCreateTT,
                AffectedSiteCreateTT,
                InternalFaultLevelCreateTT,
                IsInitialDiagnosisCreateTT,
                IsVIPCreateTT,
                IVRCallLevelCreateTT,
                ModuleTypeCreateTT,
                NetworkTypeCreateTT,
                OrderIdCreateTT,
                ParentOrderIDConfirmTT,
                FaultOccurrenceTimeCreateTT,
                ProductTypeCreateTT,
                RegionCreateTT,
                SeverityCreateTT,
                SiteOutageCreateTT,
                SiteIDCreateTT,
                SitePriorityCreateTT,
                SiteTypeCreateTT,
                SourceTicketIDCreateTT,
                TemplateCreateTT,
                SummaryCreateTT_,
                TroubleSourceCreateTT,
                CreateTimeCreateTT,
                OperatorCreateTT,
                SubmitTimeCreateTT,
                VendorCreateTT,
                SuspectedRCDescriptionHandleT,
                OperationTimeHandleTT,
                DescriptionHandleTT,
                OutageSiteQTYCreateTT_,
                OperationHandleTT,
                OrginalGroupHandleTT,
                ParentOrderIDConfirmTT_dd,
                SuspectedRCsuspectedrcProcess_,
                CreateTimeHandleTT,
                OperatorHandleTT,
                SubmitTimeHandleTT,
                DescriptionassigndescriptionP,
                DiagnoseCauseProcessTT,
                DiagnoseControlProcessTT,
                DiagnoseKeyInfoProcessTT,
                DiagnoseSuggestionProcessTT,
                DiagnoseTTFlagProcessTT,
                ETRinHoursProcessTT,
                FaultFirstOccurTimeCreateTT_,
                FaultLastOccurTimeCreateTT_,
                FaultRecoveryTimeProcessTT,
                FaultResolveOrganizationProce,
                FaultResolvingTimeProcessTT,
                WFMCMNoProcessTT,
                IsLoadDataProcessTT,
                ModuleTypeCreateTT_,
                NeedSpareProcessTT,
                NetworkTypeCreateTT_,
                OperationProcessTT,
                OperatorProcessTT,
                OriginatorProcessTT_,
                ParentOrderIDConfirmTT_d,
                ParentTicketProcessTT,
                PartnerTicketIdProcessTT,
                ReasonForOutageProcessTT,
                DescriptionprocrecorddesProce,
                AttachmentProcessTT,
                DescriptionprocupdatedesProce,
                ProductTypeCreateTT_,
                AssignToProcessTT,
                ResponsibilityProcessTT,
                ServiceInterruptionTimeProces,
                SolutionTypeProcessTT,
                SpareinformationProcessTT,
                Assignto3rdPartyOrganizationP,
                SubSolutionTypeProcessTT,
                SuspectedRCsuspectedrcProcess,
                CreateTimeProcessTT,
                OperatorProcessTT_,
                SubmitTimeProcessTT,
                WFMCompleteDesciptionResolveT_,
                WFMRCAResolveTT,
                CreateTimeResolveTT,
                OperatorResolveTT,
                SubmitTimeResolveTT,
                Level3ResolveTT,
                OperationResolveTT,
                ParentOrderIDConfirmTT__,
                POFSiteIDResolveTT,
                RejectDescriptionResolveTT,
                SolutionResolveTT,
                ReasonForHMTTRResolveTT,
                Level1ResolveTT,
                Level2ResolveTT,
                WFMCompleteDesciptionResolveT,
                WFMRCAResolveTT_,
                BODepartmentProcessTT,
                BOFirstHandleTimeProcessTT,
                BOProcessorProcessTT,
                FirstAssignBOTimeProcessTT,
                MTTP,
                MTTR,
                MTTC,
                MTT_ACK,
                LAST_UPDATE_TIME
                    ) VALUES """
    for i in range(0, len(viz)):
        if str(viz[nas[93]][i]) != 'nan' and str(viz[nas[41]][i]) != 'nan':
            d1 = datetime.datetime.strptime(str(viz[nas[41]][i]), '%Y-%m-%d %H:%M:%S')
            d2 = datetime.datetime.strptime(str(viz[nas[93]][i]), '%Y-%m-%d %H:%M:%S')
            mttr = int((d2 - d1).total_seconds() / 60)
        else:
            mttr = ''

        if str(viz[nas[9]][i]) != 'nan' and str(viz[nas[93]][i]) != 'nan':
            d1 = datetime.datetime.strptime(str(viz[nas[9]][i]), '%Y-%m-%d %H:%M:%S')
            d2 = datetime.datetime.strptime(str(viz[nas[93]][i]), '%Y-%m-%d %H:%M:%S')
            mttc = int((d1 - d2).total_seconds() / 60)
        else:
            mttc = ''

        
        if str(viz[nas[93]][i]) != 'nan' and str(viz[nas[126]][i]) != 'nan':
            d2 = datetime.datetime.strptime(str(viz[nas[93]][i]), '%Y-%m-%d %H:%M:%S')
            d1 = datetime.datetime.strptime(str(viz[nas[126]][i]), '%Y-%m-%d %H:%M:%S')
            mttp = int((d2 - d1).total_seconds() / 60)
        else:
            mttp = ''
        
        if str(viz[nas[126]][i]) != 'nan' and str(viz[nas[128]][i]) != 'nan':
            d2 = datetime.datetime.strptime(str(viz[nas[126]][i]), '%Y-%m-%d %H:%M:%S')
            d1 = datetime.datetime.strptime(str(viz[nas[128]][i]), '%Y-%m-%d %H:%M:%S')
            mtt_ack = int((d2 - d1).total_seconds() / 60)
        else:
            mtt_ack = ''
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
        viz[nas[12]][i],
        viz[nas[13]][i],
        viz[nas[14]][i],
        viz[nas[15]][i],
        viz[nas[16]][i],
        viz[nas[17]][i],
        viz[nas[18]][i],
        viz[nas[19]][i],
        viz[nas[20]][i],
        viz[nas[21]][i],
        viz[nas[22]][i],
        str(viz[nas[23]][i]).replace("'",""),
        str(viz[nas[24]][i]).replace("'",""),
        str(viz[nas[25]][i]).replace("'","").replace('"',''),
        viz[nas[26]][i],
        viz[nas[27]][i],
        viz[nas[28]][i],
        viz[nas[29]][i],
        viz[nas[30]][i],
        str(viz[nas[31]][i])[:20].replace("'",""),
        str(viz[nas[32]][i]).replace("'",""),
        viz[nas[33]][i],
        viz[nas[34]][i],
        viz[nas[35]][i],
        viz[nas[36]][i],
        viz[nas[37]][i],
        viz[nas[38]][i],
        viz[nas[39]][i],
        viz[nas[40]][i],
        viz[nas[41]][i],
        viz[nas[42]][i],
        viz[nas[43]][i],
        viz[nas[44]][i],
        viz[nas[45]][i],
        viz[nas[46]][i],
        viz[nas[47]][i],
        viz[nas[48]][i],
        viz[nas[49]][i],
        viz[nas[50]][i],
        viz[nas[51]][i],
        viz[nas[52]][i],
        viz[nas[53]][i],
        viz[nas[54]][i],
        viz[nas[55]][i],
        viz[nas[56]][i],
        viz[nas[57]][i],
        viz[nas[58]][i],
        viz[nas[59]][i],
        viz[nas[60]][i],
        viz[nas[61]][i],
        viz[nas[62]][i],
        viz[nas[63]][i],
        viz[nas[64]][i],
        viz[nas[65]][i],
        viz[nas[66]][i],
        viz[nas[67]][i],
        viz[nas[68]][i],
        viz[nas[69]][i],
        viz[nas[70]][i],
        str(viz[nas[71]][i])[:20].replace("'",""),
        viz[nas[72]][i],
        viz[nas[73]][i],
        viz[nas[74]][i],
        viz[nas[75]][i],
        viz[nas[76]][i],
        viz[nas[77]][i],
        viz[nas[78]][i],
        viz[nas[79]][i],
        viz[nas[80]][i],
        viz[nas[81]][i],
        viz[nas[82]][i],
        viz[nas[83]][i],
        str(viz[nas[84]][i])[:20].replace("'",""),
        viz[nas[85]][i],
        viz[nas[86]][i],
        viz[nas[87]][i],
        viz[nas[88]][i],
        viz[nas[89]][i],
        viz[nas[90]][i],
        viz[nas[91]][i],
        viz[nas[92]][i],
        viz[nas[93]][i],
        viz[nas[94]][i],
        viz[nas[95]][i],
        viz[nas[96]][i],
        viz[nas[97]][i],
        viz[nas[98]][i],
        viz[nas[99]][i],
        viz[nas[100]][i],
        viz[nas[101]][i],
        viz[nas[102]][i],
        viz[nas[103]][i],
        viz[nas[104]][i],
        viz[nas[105]][i],
        viz[nas[106]][i],
        viz[nas[107]][i],
        viz[nas[108]][i],
        viz[nas[109]][i],
        viz[nas[110]][i],
        viz[nas[111]][i],
        viz[nas[112]][i],
        viz[nas[113]][i],
        viz[nas[114]][i],
        viz[nas[115]][i],
        viz[nas[116]][i],
        viz[nas[117]][i],
        viz[nas[118]][i],
        viz[nas[119]][i],
        viz[nas[120]][i],
        viz[nas[121]][i],
        viz[nas[122]][i],
        viz[nas[123]][i],
        viz[nas[124]][i],
        viz[nas[129]][i],
        viz[nas[130]][i],
        viz[nas[131]][i],
        viz[nas[132]][i],
        viz[nas[129]][i],
        viz[nas[130]][i],
        viz[nas[131]][i],
        viz[nas[132]][i],
        viz[nas[133]][i],
        viz[nas[134]][i],
        viz[nas[135]][i],
        viz[nas[136]][i],
        viz[nas[137]][i],
        viz[nas[138]][i],
        viz[nas[125]][i],
        viz[nas[126]][i],
        viz[nas[127]][i],
        viz[nas[128]][i],
        mttp,
        mttr,
        mttc,
        mtt_ack,
        insert_date 
            )

        try: 
            corsor.execute(query + str(data).replace("nan","''").replace('"',"").replace("\\"," ").replace("  "," "))
            print("insert successfully")
        except Exception as e:
            log.write(str(e) + '\n')
            #print(e)
            #print(str(data).replace("nan","''").replace('"',"").replace("\\"," ").replace("  "," "))
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
        "OrderId",
        "WorkflowType",
        "SummaryCreateTT",
        "TicketStatus",
        "BusinessStatus",
        "SLAStatus",
        "CurrentOperator",
        "OriginatorProcessTT",
        "CreateTime",
        "CloseTime",
        "CurrentPhase",
        "Escalated",
        "Exempted",
        "DescriptioncloseacceptdesConf",
        "AttachmentConfirmTT",
        "DescriptioncloserejectdesConf",
        "SatisfactionDegreeConfirmTT",
        "OperationConfirmTT",
        "ParentOrderIDConfirmTT_",
        "CreateTimeConfirmTT",
        "OperatorConfirmTT",
        "SubmitTimeConfirmTT",
        "AffectedRegionCreateTT",
        "AlarmIDCreateTT",
        "AlarmNameCreateTT",
        "AlarmTypeCreateTT",
        "BSCCreateTT",
        "BSC_RNCCreateTT",
        "AssignToCreateTT",
        "AchmentCreateTT",
        "CopyToCreateTT",
        "DescriptionCreateTT",
        "SolutionCreateTT",
        "DeviceCreateTT",
        "DomainCreateTT",
        "EMSCreateTT",
        "EMSAlarmNoCreateTT",
        "EventCodeCreateTT",
        "AcknowledgementTimeCreateTT",
        "CausedByCRCreateTT",
        "FaultFirstOccurTimeCreateTT",
        "FaultLastOccurTimeCreateTT",
        "FaultLevelCreateTT",
        "FaultNoCreateTT",
        "TargetTimeCreateTT",
        "ImpactCreateTT",
        "ImpactedServiceCreateTT",
        "OutageSiteQTYCreateTT",
        "AffectedSiteCreateTT",
        "InternalFaultLevelCreateTT",
        "IsInitialDiagnosisCreateTT",
        "IsVIPCreateTT",
        "IVRCallLevelCreateTT",
        "ModuleTypeCreateTT",
        "NetworkTypeCreateTT",
        "OrderIdCreateTT",
        "ParentOrderIDConfirmTT",
        "FaultOccurrenceTimeCreateTT",
        "ProductTypeCreateTT",
        "RegionCreateTT",
        "SeverityCreateTT",
        "SiteOutageCreateTT",
        "SiteIDCreateTT",
        "SitePriorityCreateTT",
        "SiteTypeCreateTT",
        "SourceTicketIDCreateTT",
        "TemplateCreateTT",
        "SummaryCreateTT_",
        "TroubleSourceCreateTT",
        "CreateTimeCreateTT",
        "OperatorCreateTT",
        "SubmitTimeCreateTT",
        "VendorCreateTT",
        "SuspectedRCDescriptionHandleT",
        "OperationTimeHandleTT",
        "DescriptionHandleTT",
        "OutageSiteQTYCreateTT_",
        "OperationHandleTT",
        "OrginalGroupHandleTT",
        "ParentOrderIDConfirmTT_dd",
        "SuspectedRCsuspectedrcProcess_",
        "CreateTimeHandleTT",
        "OperatorHandleTT",
        "SubmitTimeHandleTT",
        "DescriptionassigndescriptionP",
        "DiagnoseCauseProcessTT",
        "DiagnoseControlProcessTT",
        "DiagnoseKeyInfoProcessTT",
        "DiagnoseSuggestionProcessTT",
        "DiagnoseTTFlagProcessTT",
        "ETRinHoursProcessTT",
        "FaultFirstOccurTimeCreateTT_",
        "FaultLastOccurTimeCreateTT_",
        "FaultRecoveryTimeProcessTT",
        "FaultResolveOrganizationProce",
        "FaultResolvingTimeProcessTT",
        "WFMCMNoProcessTT",
        "IsLoadDataProcessTT",
        "ModuleTypeCreateTT_",
        "NeedSpareProcessTT",
        "NetworkTypeCreateTT_",
        "OperationProcessTT",
        "OperatorProcessTT",
        "OriginatorProcessTT_",
        "ParentOrderIDConfirmTT_d",
        "ParentTicketProcessTT",
        "PartnerTicketIdProcessTT",
        "ReasonForOutageProcessTT",
        "DescriptionprocrecorddesProce",
        "AttachmentProcessTT",
        "DescriptionprocupdatedesProce",
        "ProductTypeCreateTT_",
        "AssignToProcessTT",
        "ResponsibilityProcessTT",
        "ServiceInterruptionTimeProces",
        "SolutionTypeProcessTT",
        "SpareinformationProcessTT",
        "Assignto3rdPartyOrganizationP",
        "SubSolutionTypeProcessTT",
        "SuspectedRCsuspectedrcProcess",
        "CreateTimeProcessTT",
        "OperatorProcessTT_",
        "SubmitTimeProcessTT",
        "WFMCompleteDesciptionResolveT_",
        "WFMRCAResolveTT",
        "CreateTimeResolveTT",
        "OperatorResolveTT",
        "SubmitTimeResolveTT",
        "Level3ResolveTT",
        "OperationResolveTT",
        "ParentOrderIDConfirmTT__",
        "POFSiteIDResolveTT",
        "RejectDescriptionResolveTT",
        "SolutionResolveTT",
        "ReasonForHMTTRResolveTT",
        "Level1ResolveTT",
        "Level2ResolveTT",
        "WFMCompleteDesciptionResolveT",
        "WFMRCAResolveTT_",
        "BODepartmentProcessTT",
        "BOFirstHandleTimeProcessTT",
        "BOProcessorProcessTT",
        "FirstAssignBOTimeProcessTT",
        "MTTP",
        "MTTR",
        "MTTC",
        "MTT_ACK",
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
    prefs = {'download.default_directory' : 'C:/Users/ewx510986/Downloads/AIRTEL_REPORT_TT'}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    weboperation(date1,date2,driver)
    return


if __name__ =='__main__':
    # log = open("C:\\Users\\ewx510986\\Desktop\\machine learning\\airtel_ows_test.txt","w")
    # _date = str(datetime.datetime.now() - datetime.timedelta(minutes=60))[:19]
    # start = _date
    # end  =  _date
    # log.write(str(datetime.datetime.now()) + '\n') 
    # main(start,end)
    data_preprocess()

    
