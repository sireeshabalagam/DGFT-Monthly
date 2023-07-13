#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
from datetime import datetime
import datetime


# In[3]:


MON_NUM=[1,2,3,4,5,6,7,8,9,10,11,12]
MON_NAME=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
Month_Source = pd.DataFrame([MON_NUM,MON_NAME]).T
Month_Source.columns =['Month Number','Month Name']
Month_Source
months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
year=['2019','2020','2021','2022','2023']


# In[7]:


links=[]
for j in year:
    for i in months:
        if (i=='APR') & (j=='2023'):
            break
        else:
            x = r"https://npp.gov.in/public-reports/cea/monthly/generation/18%20col%20act/"+j+"/"+i+"//18%20col%20act-18_"+j+'-'+i+".xls"
            links.append(x)


# In[32]:


def Data(excel):
    df = pd.read_excel(excel,header=[4]).iloc[1:-1,[1,2,3,4,5]]
    df.insert(0,'Month',df.columns[2])
    df['Month'] = ' ' + df['Month']
    df.columns = ["Month","Station","Monitered Capacity (Mega Watt)","Target","Generation (Million Unit)","Plant Load Factor Actual (In percent %)"]
    df.insert(0,'Year',df.Month.str.split('-',expand=True)[1])
    df.insert(0,'Country','India')
    return df


  


# In[33]:


final = []
for i in links[6:]:
    final.append(Data(i))


# In[34]:


final_df = pd.concat(final,ignore_index=True)


# In[35]:


final_df.to_csv('NDI-1187-401_GRID OPERATION & DISTRIBUTION WING- OPERATION PERFORMANCE MONITORING DIVISION-SECTOR WISE DETAILS_CENTRAL.csv',index=False,encoding='utf-8')


# In[ ]:




