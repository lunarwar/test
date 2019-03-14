import cx_Oracle
import pandas as pd
import datetime
from datetime import date, timedelta
import random
import datetime

def OWS(table,dates,date2): 
    exportID = str(random.randint(1,500000))  
    _date = str(datetime.datetime.today())[:10]
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    data = pd.read_sql("select * from "+ table +" where TO_DATE(CREATETIME,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and\
        TO_DATE(CREATETIME,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)
    con.close()
    export_file_name = table+"_"+exportID+"_"+_date
    #data.head() # take a peek at data
    data.to_excel(export_file_name+".xlsx", index=False)
    return export_file_name

def MAPS(table,dates,date2): 
    exportID = str(random.randint(1,500000)) 
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    if table =="2G AVAILABILITY":
        tech = "2G"
        table = "MTN_MAPS"
        data = pd.read_sql("select * from "+ table +" where technology = '"+ tech + "' and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)

    elif table =="3G AVAILABILITY":
        tech = "3G"
        table = "MTN_MAPS"
        data = pd.read_sql("select * from "+ table +" where technology = '"+ tech + "' and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)

    elif table =="4G AVAILABILITY":
        tech = "4G"
        table = "MTN_MAPS"
        data = pd.read_sql("select * from "+ table +" where technology = '"+ tech + "' and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)
    
    elif table == "2G/3G/4G":
        tech = "4G"
        table = "MTN_MAPS"
        data = pd.read_sql("select * from "+ table +" where  \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and \
        TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)
    
    else:
        ss = dates.split('-')
        ssy = date2.split('-')
        new_sdate = ss[1]+"-"+ss[2]+"-"+ss[0]
        new_sdate2 = ssy[1]+"-"+ssy[2]+"-"+ssy[0]
        data = pd.read_sql("select * from "+ table +" where TO_DATE(time,'MM/DD/YYYY,HH24:MI:SS') >= TO_DATE('" + new_sdate + "','MM/DD/YYYY,HH24:MI:SS') and\
        TO_DATE(time,'MM/DD/YYYY,HH24:MI:SS') <= TO_DATE('" +new_sdate2 + "','MM/DD/YYYY,HH24:MI:SS')", con)

    con.close()
    export_file_name = table+exportID
    #data.head() # take a peek at data
    data.to_excel(table+exportID+".xlsx", index=False)
    return export_file_name

def IRESOURCE(table,skillv,jobrole,source): 
    exportTime = str(datetime.datetime.today())[:10]  
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    data = pd.read_sql("select * from "+ table +" where  '" + skillv + "' and jobrole" , con)
    con.close()
    #data.head() # take a peek at data
    data.to_excel('testfile.xlsx')

def wimax_lte(table,dates,date2):
    _date = str(datetime.datetime.today())[:10]
    exportID = str(random.randint(1,500000)) 
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    data = pd.read_sql("select * from "+ table +" where  \
    TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + dates + "','YYYY/MM/DD,HH24:MI:SS') and \
    TO_DATE(time,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" +date2 + "','YYYY/MM/DD,HH24:MI:SS')", con)
    con.close()
    export_file_name = table+"_"+exportID+"_"+_date
    #data.head() # take a peek at data
    data.to_excel(export_file_name+".xlsx", index=False)
    return export_file_name

def worstPsites(table, dates, date2):
    _date = str(datetime.datetime.today())[:10]
    exportID = str(random.randint(1,500000)) 
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    data = pd.read_sql("SELECT \
    mtn_reporting_ows.faultlastoccurtimehwprocesstt AS FAULT_LAST_OCCURE_TIME,\
    mtn_reporting_ows.devicenamecreatett AS DEVICE_NAME, \
    mtn_maps.technology AS TECHNOLOGY,\
    mtn_maps.region AS REGION,\
    mtn_maps.TERRITORY, \
    mtn_maps.availabilityratecell AS AVAILABILITY_SITE, \
    mtn_reporting_ows.orderid,\
    mtn_reporting_ows.alarmnamecreatett AS ALARM_NAME,\
    mtn_reporting_ows.outagesiteqtycreatett AS OUTAGE_QUANTITY, \
    mtn_reporting_ows.mttr, \
    mtn_reporting_ows.rcabuckethwprocesstt AS RCA_BUCKET,\
    mtn_maps.status\
    FROM mtn_reporting_ows \
       JOIN mtn_maps ON mtn_reporting_ows.devicenamecreatett = mtn_maps.bts_nodeb\
            WHERE mtn_maps.availabilityratecell < 90 \
                AND TO_DATE(mtn_reporting_ows.CREATETIME, 'YYYY/MM/DD HH24:MI:SS') >= TO_DATE('"+dates+"', 'YYYY/MM/DD HH24:MI:SS')\
                    AND TO_DATE(mtn_reporting_ows.CREATETIME, 'YYYY/MM/DD HH24:MI:SS') < TO_DATE('"+date2+"', 'YYYY/MM/DD HH24:MI:SS')\
                        AND TO_DATE(mtn_maps.time, 'YYYY/MM/DD HH24:MI:SS') >= TO_DATE('"+dates+"', 'YYYY/MM/DD HH24:MI:SS')\
                            AND TO_DATE(mtn_maps.time, 'YYYY/MM/DD HH24:MI:SS') < TO_DATE('"+date2+"', 'YYYY/MM/DD HH24:MI:SS') \
                                order BY mtn_maps.availabilityratecell", con)
    con.close()
    export_file_name = table+"_"+exportID+"_"+_date
    #data.head() # take a peek at data
    data.to_excel(export_file_name+".xlsx",index=False)
    return export_file_name
# incident management dashboard
def imdashboard(table, dates, date2):
    _date = str(datetime.datetime.today())[:10]
    exportID = str(random.randint(1,500000)) 
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 
    data = pd.read_sql("SELECT DISTINCT (mtn_reporting_ows.devicenamecreatett), \
    mtn_reporting_ows.outagesiteqtycreatett, mtn_maps.TERRITORY,mtn_maps.bts_nodeb,mtn_maps.time,\
    mtn_reporting_ows.faultlastoccurtimehwprocesstt, mtn_maps.availabilityratecell, \
    mtn_reporting_ows.mttr,mtn_reporting_ows.ticketstatus,\
    mtn_reporting_ows.rcabuckethwprocesstt,mtn_reporting_ows.orderid,\
    mtn_reporting_ows.faultrecoverytimefaultrecovery, mtn_reporting_ows.IMPACTED4GQTYCREATETT,mtn_reporting_ows.IMPACTED3GQTYCREATETT, mtn_reporting_ows.IMPACTED2GQTYCREATETT,\
    mtn_maps.technology, mtn_maps.region, mtn_reporting_ows.FAULTREASONDESCRIPTIONHWPROCES, mtn_reporting_ows.SOLUTIONHWPROCESSTT\
  FROM mtn_reporting_ows \
       JOIN mtn_maps ON mtn_reporting_ows.devicenamecreatett = mtn_maps.bts_nodeb \
            WHERE mtn_reporting_ows.outagesiteqtycreatett >= 10 \
                AND TO_DATE(mtn_reporting_ows.CREATETIME, 'YYYY/MM/DD HH24:MI:SS') >= TO_DATE('"+ dates +"', 'YYYY/MM/DD HH24:MI:SS')\
                    AND TO_DATE(mtn_reporting_ows.CREATETIME, 'YYYY/MM/DD HH24:MI:SS') < TO_DATE('"+ date2 +"', 'YYYY/MM/DD HH24:MI:SS')\
                        AND TO_DATE(mtn_maps.time, 'YYYY/MM/DD HH24:MI:SS') >= TO_DATE('"+ dates +"', 'YYYY/MM/DD HH24:MI:SS')\
                            AND TO_DATE(mtn_maps.time, 'YYYY/MM/DD HH24:MI:SS') < TO_DATE('"+ date2 +"', 'YYYY/MM/DD HH24:MI:SS') \
                                order BY mtn_reporting_ows.devicenamecreatett", con)
    con.close()
    export_file_name = table+"_"+exportID+"_"+_date
    data.to_excel(export_file_name+".xlsx",index=False)
    return export_file_name

def ems_u2000(table, start_, end_):
    _date = str(datetime.datetime.today())[:10]
    exportID = str(random.randint(1,500000)) 
    con = cx_Oracle.connect('datacenter/datacenter@127.0.0.1/orcl') 

    data = pd.read_sql("select * from "+ table +" where region = 'ABUJA"+"' and \
        TO_DATE(OccurrenceTime,'YYYY/MM/DD,HH24:MI:SS') >= TO_DATE('" + start_ + "','YYYY/MM/DD,HH24:MI:SS') and \
        TO_DATE(OccurrenceTime,'YYYY/MM/DD,HH24:MI:SS') <= TO_DATE('" + end_ + "','YYYY/MM/DD,HH24:MI:SS')", con)

    con.close()
    export_file_name = table+"_"+exportID+"_"+_date
    data.to_excel(export_file_name+".xlsx",index=False)
    return export_file_name



    

    
