from distutils.dir_util import copy_tree
import os

copy_tree("//172.24.248.86/users/huautma/desktop/output", "C:/users/administrator/desktop/robot")
copy_tree("C:/users/administrator/desktop/robot", "//172.16.151.73/users/automation-desk/desktop/ericsson oss-rc")
print("all files copied to automation server successfully")


import os
import shutil
import datetime


def copyfiless():
    root_path ='C:/Users/automation-desk/Downloads/' 
    arr = os.listdir('C:/Users/automation-desk/Downloads/')
    app_date = str(datetime.datetime.today())[:20].replace(':','')
    for data in arr:
        # all outage alarms
        if data == 'Incident ticket dump.xlsx': #specify the starting letters of file name to be copied
            source = root_path + data # initize file
            target = 'C:/Users/automation-desk/Desktop/report automation/report automation/TT dump' # destination dir
            move_file(source, target)
            source2 = root_path + data # final file source
            os.rename(source2, root_path + app_date + data) # new file name
            print(data + "was copied successfully")
        
        if data == 'Incident ticket.xlsx': #specify the starting letters of file name to be copied
            source = root_path + data # initize file
            target = 'C:/Users/automation-desk/Desktop/report automation/report automation/Reconcilliation/input data' # destination dir
            move_file(source, target)
            rsource = root_path + data # final file source
            os.rename(rsource, root_path + app_date + data) # new file name
            print(data + "was copied successfully")

        if data == 'History Alarms_All Region.xlsx': #specify the starting letters of file name to be copied
            source = root_path + data # initize file
            target = 'C:/Users/automation-desk/Desktop/report automation/report automation/Fluctuation Alarms' # destination dir
            move_file(source, target)
            rsource2 = root_path + data # final file source
            os.rename(rsource2, root_path + app_date + data) # new file name
            print(data + "was copied successfully")

        if data == 'pending_running_tt.xlsx': #specify the starting letters of file name to be copied
            source = root_path + data # initize file
            target = 'C:/Users/automation-desk/Desktop/TT dashboard' # destination dir
            move_file(source, target)
            rsource2 = root_path + data # final file source
            os.rename(rsource2, root_path + app_date + data) # new file name
            print(data + "was copied successfully")


def move_file(source, target):        
    try:
        shutil.copy(source, target)    
    except IOError as e:
        print("Unable to copy file. %s" % e)
    except:
        print("Unexpected error:", sys.exc_info())

copyfiless()


