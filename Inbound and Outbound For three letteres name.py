#!/usr/bin/env python
# coding: utf-8

# In[54]:



import pandas as pd
import numpy as np


# In[55]:


IB =  pd.read_excel(r"C:\Users\3363\Desktop\Reports\Inbound.xlsx")
IB.columns = IB.iloc[0]
IB = IB.drop(IB.index[0])
IB = IB.iloc[:-2 , :]

IB.AGENT=IB.AGENT.str.replace('-','')
IB[['ID','Agent', 'Last', 'Late']]= IB.AGENT.str.split(expand=True)
IB = IB [['ID', 'CALLS']]
IB = IB.rename(columns = {'CALLS_x' : 'Inbound', 'CALLS_y': 'Outbound'})


# In[56]:


IB


# In[57]:


OB =  pd.read_excel(r"C:\Users\3363\Desktop\Reports\outbound.xlsx")

OB.AGENT=OB.AGENT.str.replace('-','')
OB[['ID','Agent', 'Last', 'Late']] = OB.AGENT.str.split(expand=True)
OB = OB [['ID', 'CALLS']]

OB.head()


# In[58]:


InOb = pd.merge(IB, OB, how = 'outer', on = 'ID')
InOb = InOb.rename(columns = {'CALLS_x' : 'Inbound', 'CALLS_y': 'Outbound'})
InOb["Inbound"].fillna(0, inplace = True)
InOb["Outbound"].fillna(0, inplace = True)
InOb['Total Calls'] = InOb['Inbound'] + InOb['Outbound']
InOb.head()


# In[59]:


InOb.to_excel(r"C:\Users\3363\Desktop\Akshay\IB Report.xlsx")


# In[ ]:




