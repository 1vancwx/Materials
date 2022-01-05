#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client
import sys
import subprocess
import time
import os

from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=105)
fisy = fis.strftime('%Y')

# This function will Login to SAP from the SAP Logon window

def saplogin(t_code):

    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(4)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection("ECP [Everest ECC Production]", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "10299976"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "password"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize
        
        if t_code == 3000 or t_code == "3000":
            session.findById("wnd[0]/tbar[0]/okcd").text = "faglb03"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtRACCT-LOW").text = "300000"
            session.findById("wnd[0]/usr/ctxtRACCT-HIGH").text = "324001"
            session.findById("wnd[0]/usr/ctxtRBUKRS-LOW").text = "3000"
            session.findById("wnd[0]/usr/txtRYEAR").text = fisy
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)
            session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell(17,"BALANCE")
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(2)
            
            session.findById("wnd[0]").sendVKey(33)
            session.findById("wnd[1]").sendVKey(71)
            time.sleep(1)
            
            session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/LM MTH Rpt"
            session.findById("wnd[2]/usr/txtSCAN_STRING-LIMIT").text = "999"
            session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = "false"
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[3]").sendVKey(84)
            session.findById("wnd[3]").sendVKey(2)
            session.findById("wnd[1]").sendVKey(0)
            
            session.findById("wnd[0]").sendVKey(38)
            session.findById("wnd[1]/usr/btnB_SEARCH").press()
            session.findById("wnd[2]/usr/txtGD_SEARCHSTR").text = "TRADING"
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/btnAPP_WL_SING").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "sg01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "My01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "vn01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "th01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "mm01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "id01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "ph01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "pk01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").setFocus
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 4
            session.findById("wnd[2]").sendVKey(82)
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "z661"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "kh01"
            session.findById("wnd[2]/tbar[0]/btn[8]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
            session.findById("wnd[0]").sendVKey(16)
                              
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\Gross to net revenue\Data"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fisy + " GL_3000.xlsx"
            session.findById("wnd[1]").sendVKey(11)
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe") 
            
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nfaglb03"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtRACCT-LOW").text = "300000"
            session.findById("wnd[0]/usr/ctxtRACCT-HIGH").text = "324001"
            session.findById("wnd[0]/usr/ctxtRBUKRS-LOW").text = "3505"
            session.findById("wnd[0]/usr/txtRYEAR").text = fisy
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)
            session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell(17,"BALANCE")
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(2)
            
          
            session.findById("wnd[0]").sendVKey(33)
            session.findById("wnd[1]").sendVKey(71)
            time.sleep(1)
            
            session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/LM MTH Rpt"
            session.findById("wnd[2]/usr/txtSCAN_STRING-LIMIT").text = "999"
            session.findById("wnd[2]/usr/chkSCAN_STRING-START").selected = "false"
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[3]").sendVKey(84)
            session.findById("wnd[3]").sendVKey(2)
            session.findById("wnd[1]").sendVKey(0)
            
            session.findById("wnd[0]").sendVKey(38)
            session.findById("wnd[1]/usr/btnB_SEARCH").press()
            session.findById("wnd[2]/usr/txtGD_SEARCHSTR").text = "TRADING"
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/btnAPP_WL_SING").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "sg01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "My01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "vn01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "th01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "mm01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "id01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "ph01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "pk01"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").setFocus
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 4
            session.findById("wnd[2]").sendVKey(82)
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "z661"
            session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "kh01"
            session.findById("wnd[2]/tbar[0]/btn[8]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            session.findById("wnd[0]").sendVKey(16)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\Gross to net revenue\Data"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fisy + " GL_3505.xlsx"
            session.findById("wnd[1]").sendVKey(11)
            time.sleep(2)
            os.system("TASKKILL /F /IM saplogon.exe") 
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe") 
    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


# In[2]:


saplogin(3000)

# In[5]:


import pandas as pd
import numpy as np
from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=105)
fisy = fis.strftime('%Y')
pth3000 = r"K:\FIN\Ivan\Gross to net revenue\Data\\" + fisy + " GL_3000.xlsx"
pth3505 = r"K:\FIN\Ivan\Gross to net revenue\Data\\" + fisy + " GL_3505.xlsx"
main_df = pd.read_excel(pth3000, dtype=str)
isa_df = pd.read_excel(pth3505, dtype=str)
cmb_df = pd.concat([main_df, isa_df], axis=0)
cmb_df = cmb_df.dropna(subset=['Company Code'])
reg_df = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Region", dtype=str)
cmb_df1 = cmb_df.merge(reg_df, how='left', on="Trading Part.BA")
pfc = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center", dtype=str)
pfc = pd.DataFrame(pfc, columns=['ECC Profit Center','Business Unit.1','Level 1','Level 1 Description'])
pfc = pfc.rename(columns ={'ECC Profit Center':'Profit Center'})
pfc = pfc.astype(str)
cmb_df2 = cmb_df1.merge(pfc, how='left', on='Profit Center')
coa = pd.read_excel(r"K:\FIN\Ivan\Gross to net revenue\Mapping\DLL COA.xlsx", dtype=str)
coa = pd.DataFrame(coa, columns=['Account','G/L Acct Long Text','HFM Group Account Description', 'GL Group'])
coa = coa.drop_duplicates()
df = cmb_df2.merge(coa, how='left', on='Account')
                       
df['Purpose'] = np.where(df['Account'] == '300000', "Gross Sales", "NaN")
df['Purpose'] = np.where((df['Account'] == '302000') & (df['Purpose'] == "NaN"), "I/Co Sales", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '302999') & (df['Purpose'] == "NaN"), "I/Co Sales", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '300003') & (df['Purpose'] == "NaN"), "Gross Sales (Manual)",df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '310002') & (df['Purpose'] == "NaN"), "FX", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '310016') & (df['Purpose'] == "NaN"), "Others", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '310001') & (df['Purpose'] == "NaN"), "E&O", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '300019') & (df['Purpose'] == "NaN"), "Handling", df['Purpose'])
df['Purpose'] = np.where((df['Account'] == '310003') & (df['Purpose'] == "NaN"), "Promotional Items", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("E&O", case=False)) & (df['Purpose'] == "NaN"), "E&O", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("FOC", case=False)) & (df['Purpose'] == "NaN"), "Promotional Items", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Margin", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Price", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Rebate", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Disc", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Subsid", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Clawback", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Gribbles", case=False)) & (df['Purpose'] == "NaN"), "Promotional Items", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Sample", case=False)) & (df['Purpose'] == "NaN"), "Promotional Items", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Free", case=False)) & (df['Purpose'] == "NaN"), "Promotional Items", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("FX", case=False)) & (df['Purpose'] == "NaN"), "FX", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("SCH", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Stock", case=False)) & (df['Purpose'] == "NaN"), "E&O", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Expire", case=False)) & (df['Purpose'] == "NaN"), "E&O", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Good", case=False)) & (df['Purpose'] == "NaN"), "E&O", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Diff", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Credit Note", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("Reverse", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Text'].str.contains("PD", case=False)) & (df['Purpose'] == "NaN"), "DM", df['Purpose'])
df['Purpose'] = np.where((df['Purpose'] == "NaN"), "Others", df['Purpose'])

wb = "K:\FIN\Ivan\Gross to net revenue\Data\Total_" + fisy + "_GL.xlsx"
with pd.ExcelWriter(wb) as ew:
    df.to_excel(ew, index=False)


# In[ ]:




