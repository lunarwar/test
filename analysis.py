import pandas as pd
import matplotlib
import matplotlib.pyplot as plt


def openxlsx(count_index):
    data = pd.read_excel('iccticket.xlsx')
    total = 0
    serv = data.join(data['Current Operator'].str.split(';', 5, expand=True).rename(columns={0:'group one', 1:'group two', 2:'group three', 3:'group four', 4:'group five'}))
    serve = serv.loc[serv['Ticket Status'] == 'Running'] # filter by Ticket Status == Running
    countlenght = len(serve.index)
    print(data)
    if count_index in list(serve['group one']):
        total += serve.groupby(['group one']).size()[count_index]
    
    if count_index in list(serve['group two']):
        total += serve.groupby(['group two']).size()[count_index] 
    
    if count_index in list(serve['group three']):
        total += serve.groupby(['group three']).size()[count_index] 
    
    if count_index in list(serve['group four']):
        total += serve.groupby(['group four']).size()[count_index] 

    if count_index in list(serve['group five']):
        total += serve.groupby(['group five']).size()[count_index] 

    return total
    

group_fo = openxlsx('group:FO')  
group_fo_sl = openxlsx('group:FO SL') 
group_mtn_fo = openxlsx('group:MTN FO')
group_sdm = openxlsx('group:SDMAdmin')
ihs_kano = openxlsx('group:I.H.S_KANO')  
ihs_lagos = openxlsx('group:I.H.S_LAGOS') 
ihs_ibadan = openxlsx('group:I.H.S_IBADAN')  
ihs_abuja = openxlsx('group:I.H.S_ABUJA') 
ihs_phc = openxlsx('group:I.H.S_PHC')
data_frame_data = {
    'OPERATORS': ['FO','FO SL','MTN FO', 'SDM ADMIN','IHS_ABUJA','IHS_KANO','IHS_LAGOS','IHS_IBADAN','IHS_PHC', 'USERS'],
    'TICKETS': [group_fo, group_fo_sl,group_mtn_fo, group_sdm, ihs_abuja, ihs_kano, ihs_lagos, ihs_ibadan, ihs_phc, 47]
    
}
new_frame = pd.DataFrame(data = data_frame_data)
new_frame.plot.bar(x='OPERATORS', y='TICKETS', rot=0)
plt.show()

