#!/usr/bin/env python
# coding: utf-8

# In[73]:



import pandas as pd
import numpy as np


# In[74]:



apr1 = pd.read_excel(r"C:\Users\3363\Desktop\Reports\one.xlsx")
apr1 = apr1.iloc[3:]
#apr1.apr1[:-1,:]
apr1 = apr1.rename(columns=apr1.iloc[0]).drop(apr1.index[0])
apr1 = apr1.iloc[:-1 , :]
apr1.head()


# In[75]:


apr2 = pd.read_excel(r"C:\Users\3363\Desktop\Reports\two.xlsx")

apr2 = apr2.iloc[3:]
#apr2.iloc[:-1,:]

apr2 = apr2.rename(columns=apr2.iloc[0]).drop(apr2.index[0])

apr2 = apr2.iloc[:-1 , :]

apr2.head()


# In[76]:


Result = pd.merge(apr1, apr2, how = 'outer', on = 'ID')
Result.head()


# In[77]:


Result1 = Result[['ID', 'USER NAME_x', 'PAUSE_x', 'WAIT', 'TALK', 'DISPO', 'DEAD', 'TOTAL', 'Idle Time','AB','LB_y','SB', 'TB', 'WB']]
Result1 = Result1.rename(columns={'USER NAME_x' : 'USER NAME', 'PAUSE_x' : 'PAUSE Time', 'LB_y': 'LB- Lunch Break', 'SB' : 'SB- Snack Break', 'TB': 'TB- Tea break', 'WB' : 'WB- Washroom break'})
#Result = Result[['ID', 'USER NAME_x','CURRENT USER GROUP_x','MOST RECENT USER GROUP_x','TOTAL','TALK', 'WAIT',  'DISPO', 'DEAD', 'Idle Time']]
#Result1.columns.values[8] = 'Idle Time'
Result1.fillna('00:00:00', inplace = True)

Result1.columns
Result1.head()


# In[78]:



Result1['Lhour'] = pd.to_datetime(Result1['LB- Lunch Break'], format='%H:%M:%S').dt.hour
Result1['Lminutes'] = pd.to_datetime(Result1['LB- Lunch Break'], format='%H:%M:%S').dt.minute
Result1['Lsecond'] = pd.to_datetime(Result1['LB- Lunch Break'], format='%H:%M:%S').dt.second


Result1['Lhour'] = Result1['Lhour'] * 3600
Result1['Lminutes'] = Result1['Lminutes'] * 60

Result1['Llunch']  = Result1['Lhour'] + Result1['Lminutes'] + Result1['Lsecond']


Result1['Thour'] = pd.to_datetime(Result1['TB- Tea break'], format='%H:%M:%S').dt.hour
Result1['Tminutes'] = pd.to_datetime(Result1['TB- Tea break'], format='%H:%M:%S').dt.minute
Result1['Tsecond'] = pd.to_datetime(Result1['TB- Tea break'], format='%H:%M:%S').dt.second



Result1['Thour']= Result1['Thour'] * 3600
Result1['Tminutes'] = Result1['Tminutes'] * 60

Result1['Ttea']  = Result1['Thour'] +Result1['Tminutes'] + Result1['Tsecond']



Result1['Whour'] = pd.to_datetime(Result1['WB- Washroom break'], format='%H:%M:%S').dt.hour
Result1['Wminutes'] = pd.to_datetime(Result1['WB- Washroom break'], format='%H:%M:%S').dt.minute
Result1['Wsecond'] = pd.to_datetime(Result1['WB- Washroom break'], format='%H:%M:%S').dt.second



Result1['Whour']= Result1['Whour'] * 3600
Result1['Wminutes'] = Result1['Wminutes'] * 60

Result1['Twb']  = Result1['Whour'] +Result1['Wminutes'] + Result1['Wsecond']

Result1['BREAK TIME1'] = Result1['Llunch'] + Result1['Ttea'] + Result1['Twb']


import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')

Result1['BREAK TIME'] = pd.to_datetime(Result1['BREAK TIME1'], unit='s').map(fmt)



# In[79]:



Result1['Ihour'] = pd.to_datetime(Result1['Idle Time'], format='%H:%M:%S').dt.hour
Result1['Iminutes'] = pd.to_datetime(Result1['Idle Time'], format='%H:%M:%S').dt.minute
Result1['Isecond'] = pd.to_datetime(Result1['Idle Time'], format='%H:%M:%S').dt.second




Result1['Ihour'] = Result1['Ihour'] * 3600
Result1['Iminutes'] = Result1['Iminutes'] * 60

Result1['IIDL']  = Result1['Ihour'] + Result1['Iminutes'] + Result1['Isecond']




Result1['ABhour'] = pd.to_datetime(Result1['AB'], format='%H:%M:%S').dt.hour
Result1['ABminutes'] = pd.to_datetime(Result1['AB'], format='%H:%M:%S').dt.minute
Result1['ABsecond'] = pd.to_datetime(Result1['AB'], format='%H:%M:%S').dt.second




Result1['ABhour'] = Result1['ABhour'] * 3600
Result1['ABminutes'] = Result1['ABminutes'] * 60

Result1['ABT']  = Result1['ABhour'] + Result1['ABminutes'] + Result1['ABsecond']


Result1['HVT BOOKING TIME1'] = Result1['IIDL'] + Result1['ABT'] 




import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')
Result1['HVT BOOKING TIME']  = pd.to_datetime(Result1['HVT BOOKING TIME1'] , unit='s').map(fmt)


# In[80]:



Result1['Tahour'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.hour
Result1['Taminutes'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.minute
Result1['Tasecond'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.second




Result1['Tahour'] = Result1['Tahour'] * 3600
Result1['Taminutes'] = Result1['Taminutes'] * 60

Result1['Tatalk']  = Result1['Tahour'] + Result1['Taminutes'] + Result1['Tasecond']




Result1['Wahour'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.hour
Result1['Waminutes'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.minute
Result1['Wasecond'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.second




Result1['Wahour'] = Result1['Wahour'] * 3600
Result1['Waminutes'] = Result1['Waminutes'] * 60

Result1['Wawt']  = Result1['Wahour'] + Result1['Waminutes'] + Result1['Wasecond']



Result1['ProductTime'] = Result1['Tatalk'] + Result1['Wawt'] + Result1['HVT BOOKING TIME1'] 






import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')
Result1['TOTAL PROD. TIME']  = pd.to_datetime(Result1['ProductTime'] , unit='s').map(fmt)


# In[81]:



Result1['TLLhour'] = pd.to_datetime(Result1['TOTAL'], format='%H:%M:%S').dt.hour
Result1['TLLminutes'] = pd.to_datetime(Result1['TOTAL'], format='%H:%M:%S').dt.minute
Result1['TLLsecond'] = pd.to_datetime(Result1['TOTAL'], format='%H:%M:%S').dt.second




Result1['TLLhour'] = Result1['TLLhour'] * 3600
Result1['TLLminutes'] = Result1['TLLminutes'] * 60


Result1['TLogin']  = Result1['TLLhour'] + Result1['TLLminutes'] + Result1['TLLsecond']


# In[82]:


Result1['Utilization %'] = Result1['ProductTime'] / Result1['TLogin']
Result1['Utilization %'] = Result1['Utilization %'].round(2)
Result1['Utilization %'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Result1['Utilization %']], index = Result1.index)


Result1.head()


# In[83]:


Result1 = Result1.drop(columns=['Lhour', 'Lminutes', 'Lsecond', 'Llunch', 'Thour', 'Tminutes', 'TLLhour','TLLminutes','TLLsecond','Tsecond', 'Whour', 'Wminutes', 'Wsecond', 'Twb','Ttea', 'Llunch', 'Ihour', 'Iminutes', 'Isecond', 'IIDL', 'ABhour', 'ABminutes', 'ABsecond', 'ABT', 'Tahour', 'Taminutes', 'Tasecond', 'Tatalk', 'Wahour', 'Waminutes', 'Wasecond','Wawt'])

Result1 = Result1.rename(columns={'AB' : 'AB HVT Booking', 'TOTAL' : 'TOTAL LOG IN TIME'})

#Result1.fillna(0, inplace = True)
#Result1['Utilization %'].fillna(0, inplace = True)


Result1.head()


# In[84]:


Result1.to_excel(r"C:\Users\3363\Desktop\Akshay\APR Report.xlsx")


# In[ ]:




