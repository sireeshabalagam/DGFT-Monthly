#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
from datetime import datetime
import datetime


# In[2]:


MON_NUM=[1,2,3,4,5,6,7,8,9,10,11,12]
MON_NAME=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
Month_Source = pd.DataFrame([MON_NUM,MON_NAME]).T
Month_Source.columns =['Month Number','Month Name']
Month_Source
months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
year=['2019','2020','2021','2022','2023']


# In[319]:


links=[]
for j in year:
    for i in months:
        if (i=='APR') & (j=='2023'):
            break
        else:
            x = r"https://npp.gov.in/public-reports/cea/monthly/generation/18%20col%20act/"+j+"/"+i+"//18%20col%20act-15_"+j+'-'+i+".xls"
            links.append(x)


# In[424]:


def Data(excel):
    df = pd.read_excel(excel,header=[7],skipfooter=8,skiprows=[8,9,10,11],usecols='A:E,N,O')
    mon=df.columns[3]
    df.columns=['Sation','Monitered Capacity (Mega Watt)','Target (Yearly)','Generation Program (Giga Watt Hour)','Generation Actual (Giga Watt Hour)','Plant Load Factor Program (In percent %)','Plant Load Factor Actual (In percent %)']


    #Sector wise indexes
    sec_ind = df[df[df.columns[1]].isnull()].index

    #for state sector 
    X = df.iloc[sec_ind[0]:sec_ind[1]]
    X.insert(0,'Sector',X[X.columns[0]].iloc[0])

    #for private sector
    Y = df.iloc[sec_ind[1]:]
    Y.insert(0,'Sector',Y[Y.columns[0]].iloc[0])


    df1 = pd.concat([X,Y],ignore_index=True)
    df1.dropna(thresh=3,inplace=True)

    df1.insert(0,'Month',mon)
    df1['Month'] = ' ' + df1['Month']
    df1.insert(0,'Year',df1.Month.str.split('-',expand=True)[1])

    df1.insert(0,'Country','India')
    return df1


# In[425]:


final = []
for i in links[6:]:
    final.append(Data(i))


# In[426]:


final_df = pd.concat(final,ignore_index=True)


# In[427]:


#df[df[df.columns[2]].isnull()].index


# In[429]:


final_df.fillna('',inplace=True)


# In[430]:


final_df[final_df.columns[10]].unique()


# In[432]:


final_df.to_csv('NDI-1184-398_GRID OPERATION & DISTRIBUTION WING-OPERATION PERFORMANCE MONITORING DIVISION-ENERGYWISE-PERFORMANCE STATUS ALL INDIA-REGIONWISE-FUEL WISE DETAILS-DIESEL.csv',index=False,encoding='utf-8')


# In[ ]:





# In[ ]:




