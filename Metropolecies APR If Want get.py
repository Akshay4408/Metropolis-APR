#!/usr/bin/env python
# coding: utf-8

# In[61]:



import pandas as pd
import numpy as np


# In[62]:


Sale = pd.read_excel(r"C:\Users\3363\Desktop\Reports\Book2.xlsx")

Sale.head()


# In[63]:


Agent = pd.read_excel(r"C:\Users\3363\Desktop\Reports\Sale Agent.xlsx")
Agent = Agent[['ID', 'USER NAME']]
Agent.head()


# In[64]:


Sale = Sale.rename(columns = {'ADVISOR NAME' : 'USER NAME'})


# In[65]:


Result = pd.merge(Sale, Agent, how = 'inner', on = 'USER NAME')
Result.head()


# In[66]:


#Sale = Sale.rename(columns = {'ADVISOR NAME' : 'USER NAME'})
Count =pd.pivot_table(data= Result, index = ['ID'], values = 'GROSS REVENUE', aggfunc = np.size, fill_value=0)
Count = Count.rename(columns = {'GROSS REVENUE' : 'Leads generated'})

Count.head()


# In[67]:


#Sale = Sale.rename(columns = {'ADVISOR NAME' : 'USER NAME'})
summ =pd.pivot_table(data= Result, index = ['ID'], values = 'GROSS REVENUE', aggfunc = np.sum, fill_value=0)
summ = summ.rename(columns = {'GROSS REVENUE' : 'Gross Revenues'})

summ.head()


# In[68]:


MainSale = pd.merge(Count, summ, how = 'outer', on = 'ID')
MainSale


# In[69]:


APR = pd.read_excel(r"C:\Users\3363\Desktop\Akshay\test.xlsx")
APR = APR.drop(columns=['Unnamed: 0'])

APR.head()


# In[70]:


APRReport = pd.merge(APR, MainSale, how = 'outer', on = 'ID')
APRReport


# In[71]:


APRReport = APRReport[['ID', 'USER NAME','Leads generated','Gross Revenues','Inbound', 'Outbound','Total Calls', 'TOTAL LOG IN TIME','TALK', 'WAIT', 'PAUSE Time', 'DISPO', 'DEAD', 'Idle Time', 'AB HVT Booking', 'HVT BOOKING TIME', 'BREAK TIME', 'LB- Lunch Break',  'SB- Snack Break', 'TB- Tea break', 'WB- Washroom break', 'TOTAL PROD. TIME', 'OCC %','Utilization %', 'AHT']]

APRReport["Leads generated"].fillna(0, inplace = True)
APRReport["Gross Revenues"].fillna(0, inplace = True)


# In[72]:


summary = pd.read_excel(r"C:\Users\3363\Desktop\Reports\APR Summary.xlsx")

summary.head()


# In[ ]:





# In[73]:




writer = pd.ExcelWriter(r"C:\Users\3363\Desktop\Akshay\Main APR Report.xlsx", engine='xlsxwriter')


# In[74]:


#APRReport.to_excel(r"C:\Users\3363\Desktop\Akshay\Main APR Report.xlsx",  index = False)


# In[75]:




APRReport.to_excel(writer, sheet_name='APRReport', index=False)
summary.to_excel(writer, sheet_name='summary')




writer.save()


# In[ ]:




