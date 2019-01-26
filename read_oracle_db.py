import cx_Oracle

def dbconnect():
    con = cx_Oracle.connect('dbusername/dbpassword@dbserver/dbservicename')
    return con


def read_table():
    con = dbconnect()
    corsor = con.cursor()
    query = """select * from mtmmmn_sdt_hourly"""
    c = corsor.execute(query)
    for row in c:
        print(row[0], "-", row[1],"-",row[2], "-", row[3],"-",row[4], "-", row[5])
    con.close()

read_table()
