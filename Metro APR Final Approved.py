#!/usr/bin/env python
# coding: utf-8

# In[48]:



import pandas as pd
import numpy as np


# In[49]:


Report = pd.read_excel(r"C:\Users\3363\Desktop\Akshay\APR Report.xlsx")
Report = Report.drop(columns=['Unnamed: 0'])

Report.head()


# In[50]:


IO = pd.read_excel(r"C:\Users\3363\Desktop\Akshay\IB Report.xlsx")
IO = IO.drop(columns=['Unnamed: 0'])

IO.head()



# In[51]:


Report = pd.merge(Report, IO, how = 'outer', on = 'ID')
Report.head()


# In[52]:


#Report = Report[['ID', 'USER NAME', 'Lead generated', 'Gross Revenues', 'Inbound', 'Outbound', 'Total Calls', 'TOTAL LOG IN TIME', 'TALK TIME', 'WAIT', 'PAUSE Time', 'DISPO', 'Idle Time', 'AB HVT Booking', 'HVT BOOKING TIME', 'DEAD TIME', 'BREAK TIME', 'LB- Lunch Break', 'SB- Snack Break', 'TB- Tea break', 'WB- Washroom break', 'ProductTime','TOTAL PROD. TIME', 'Utilization %']]
Report['AHT1'] = Report['ProductTime'] / Report['Total Calls']

Report["AHT1"].fillna(0, inplace = True)



import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')
Report['AHT'] = pd.to_datetime(Report['AHT1'] , unit='s').map(fmt)

Report.head()


# In[53]:


Report = Report.drop(columns=['TLogin', 'AHT1', 'ProductTime'])


# In[54]:


Report.columns


# In[55]:


Report = Report[['ID', 'USER NAME','Inbound', 'Outbound','Total Calls', 'TOTAL LOG IN TIME','TALK', 'WAIT', 'PAUSE Time', 'DISPO', 'DEAD', 'Idle Time', 'AB HVT Booking', 'HVT BOOKING TIME', 'BREAK TIME', 'LB- Lunch Break',  'SB- Snack Break', 'TB- Tea break', 'WB- Washroom break', 'TOTAL PROD. TIME', 'Utilization %', 'AHT']]


# In[56]:


Report["Inbound"].fillna(0, inplace = True)
Report["Outbound"].fillna(0, inplace = True)
Report["Total Calls"].fillna(0, inplace = True)

Report.fillna('00:00:00', inplace = True)


# In[57]:



Report['TLLhour'] = pd.to_datetime(Report['TOTAL LOG IN TIME'], format='%H:%M:%S').dt.hour
Report['TLLminutes'] = pd.to_datetime(Report['TOTAL LOG IN TIME'], format='%H:%M:%S').dt.minute
Report['TLLsecond'] = pd.to_datetime(Report['TOTAL LOG IN TIME'], format='%H:%M:%S').dt.second




Report['TLLhour'] = Report['TLLhour'] * 3600
Report['TLLminutes'] = Report['TLLminutes'] * 60


Report['TLogin']  = Report['TLLhour'] + Report['TLLminutes'] + Report['TLLsecond']





Report['TPLhour'] = pd.to_datetime(Report['TOTAL PROD. TIME'], format='%H:%M:%S').dt.hour
Report['TPLminutes'] = pd.to_datetime(Report['TOTAL PROD. TIME'], format='%H:%M:%S').dt.minute
Report['TPLsecond'] = pd.to_datetime(Report['TOTAL PROD. TIME'], format='%H:%M:%S').dt.second




Report['TPLhour'] = Report['TPLhour'] * 3600
Report['TPLminutes'] = Report['TPLminutes'] * 60


Report['TPProduct']  = Report['TPLhour'] + Report['TPLminutes'] + Report['TPLsecond']



Report['TBLhour'] = pd.to_datetime(Report['BREAK TIME'], format='%H:%M:%S').dt.hour
Report['TBLminutes'] = pd.to_datetime(Report['BREAK TIME'], format='%H:%M:%S').dt.minute
Report['TBLsecond'] = pd.to_datetime(Report['BREAK TIME'], format='%H:%M:%S').dt.second




Report['TBLhour'] = Report['TBLhour'] * 3600
Report['TBLminutes'] = Report['TBLminutes'] * 60


Report['TBreak']  = Report['TBLhour'] + Report['TBLminutes'] + Report['TBLsecond']






Report['TDLhour'] = pd.to_datetime(Report['DEAD'], format='%H:%M:%S').dt.hour
Report['TDLminutes'] = pd.to_datetime(Report['DEAD'], format='%H:%M:%S').dt.minute
Report['TDLsecond'] = pd.to_datetime(Report['DEAD'], format='%H:%M:%S').dt.second




Report['TDLhour'] = Report['TDLhour'] * 3600
Report['TDLminutes'] = Report['TDLminutes'] * 60


Report['TDead']  = Report['TDLhour'] + Report['TDLminutes'] + Report['TDLsecond']



Report['OCC %'] = (Report['TDead'] + Report['TBreak'] + Report['TPProduct']) / Report['TLogin']





Report['OCC %'] = Report['OCC %'].round(2)
Report['OCC %'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Report['OCC %']], index = Report.index)
Report.head()
                


# In[58]:


Report = Report[['ID', 'USER NAME','Inbound', 'Outbound','Total Calls', 'TOTAL LOG IN TIME','TALK', 'WAIT', 'PAUSE Time', 'DISPO', 'DEAD', 'Idle Time', 'AB HVT Booking', 'HVT BOOKING TIME', 'BREAK TIME', 'LB- Lunch Break',  'SB- Snack Break', 'TB- Tea break', 'WB- Washroom break', 'TOTAL PROD. TIME', 'OCC %','Utilization %', 'AHT']]


# In[59]:


Report.to_excel(r"C:\Users\3363\Desktop\Akshay\test.xlsx")

