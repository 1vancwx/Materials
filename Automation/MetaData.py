#!/usr/bin/env python
# coding: utf-8

# In[1]:


import subprocess
import os
import fnmatch
import pandas as pd

# subprocess.call(r'net use z: /del /Y', shell=True)
subprocess.call(r'net use z: https://collaborate.bdx.com/sites/bpi/ffa/MRA/Shared%20Documents /user:ivan.cho@bd.com Vo2195pr@9112', shell=True)
# os.system(r'net use z: https://collaborate.bdx.com/sites/bpi/ffa/MRA/Shared%20Documents \user:ivan.cho@bd.com Vo2195pr@91')
path = "z:/"
dir_list = os.listdir(path)

for file in os.listdir(path):
    if fnmatch.fnmatch(file, "*product hier*"):
        fpath = path + file
#         df1 = pd.read_excel(fpath)
#         print(df1[1000:1020])
#         print(fpath)


# In[2]:


df1 = pd.read_excel(fpath,"Table", usecols= "G:J,O:X", skiprows=14)
a = ['ECC6','ECC4','ACTF','LOTS','IND4','THBL'] 
df1 = df1[df1['Logical System'].isin(a)]
df1


# In[3]:


mm = df1[~df1['Reltio Product'].str.contains("PC_")]
mm['Concatenate']= mm['Business Unit.1'] + mm['Reltio Product']
mm


# In[4]:


pc= df1[df1['Reltio Product'].str.contains("PC_")]
pc['Profit Center'] = pc['Reltio Product'].str[-5:]
pc = pc.rename(columns ={'Profit Center':'ECC Profit Center'})
# pc[100:120]


# In[5]:


ids = pd.DataFrame({'PH1':['2109003','2109001','2109004','2109005','2111004','2111001','2111006','2118000','2111003','2111460','2211005','2111000','2111005','2109002']})
ids['Alt WWB'] = ['PAS','PAS','PAS','PAS','DS','DS','DS','DS','DS','DS','DS','DS','DS','PAS']
reg = pd.DataFrame({'Trading Part.BA':['TH01','SG01','VN01','MY01','ID01','PH01','KH01','Z661','MM01','PK01']})
reg['Region'] = ['SEA','SEA','SEA','SEA','SEA','SEA','OTA','OTA','OTA','Pakistan']


# In[7]:


wb = "K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx"
with pd.ExcelWriter(wb) as ew:
    mm.to_excel(ew, index=False, sheet_name='Mat Map')
    pc.to_excel(ew, index=False, sheet_name='Profit Center')
    ids.to_excel(ew, index=False, sheet_name='IDS Map')
    reg.to_excel(ew, index=False, sheet_name='Region')
 

# In[ ]:




