#!/usr/bin/env python
# coding: utf-8

# In[24]:



import pandas as pd
import numpy as np


# In[25]:


IB =  pd.read_excel(r"C:\Users\3363\Desktop\Reports\Inbound.xlsx")
IB.columns = IB.iloc[0]
IB = IB.drop(IB.index[0])

IB = IB.iloc[:-2 , :]

#IB[['ID','Agent']] = IB.AGENT.str.split(expand=True)


IB.AGENT=IB.AGENT.str.replace('-','')
IB[['ID','Agent', 'Last']] = IB.AGENT.str.split(expand=True)
IB = IB [['ID', 'CALLS']]
IB = IB.rename(columns = {'CALLS_x' : 'Inbound', 'CALLS_y': 'Outbound'})

IB.head()

IB.head()


# In[ ]:


OB =  pd.read_excel(r"C:\Users\3363\Desktop\Reports\outbound.xlsx")

OB.AGENT=OB.AGENT.str.replace('-','')
OB[['ID','Agent', 'Last']] = OB.AGENT.str.split(expand=True)
OB = OB [['ID', 'CALLS']]

OB.head()


# In[ ]:


InOb = pd.merge(IB, OB, how = 'outer', on = 'ID')
InOb = InOb.rename(columns = {'CALLS_x' : 'Inbound', 'CALLS_y': 'Outbound'})
InOb["Inbound"].fillna(0, inplace = True)
InOb["Outbound"].fillna(0, inplace = True)
InOb['Total Calls'] = InOb['Inbound'] + InOb['Outbound']
InOb.head()


# In[ ]:


#Result.info()


# In[ ]:


InOb.to_excel(r"C:\Users\3363\Desktop\Akshay\IB Report.xlsx")


# In[ ]:




