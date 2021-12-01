#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client
import sys
import subprocess
import time
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta 
from datetime import date
from dateutil.relativedelta import relativedelta
mtod = datetime.now() + relativedelta(days=5)
tod = datetime.today()
tdy = datetime.today().strftime('%d_%m_%Y')
ttdy = tod.strftime("%I%p")

if date.today().day >= 15 :
    tprd = date.today().month + 3
else:
    tprd = date.today().month + 2

prd = str(tprd).format('%#m')
fprd = str(tprd).format('%m')

if date.today().month >= 10 :
    fis = datetime.today().year + 1
else:
    fis = datetime.today().year
    
fisy = str(fis).format('%Y')
sfisy = str(fis).format('%y')
sprd = (datetime.today()-relativedelta(days=10)).strftime('%b')

holpath = r"\\apsgtusan01\sgtu\Dept\sgseafin\FIN\Ivan\Gross to net revenue\BotInput\BotScheduler.xlsx"
hols = pd.read_excel(holpath,sheet_name="Holiday",header=None)
hols = pd.DataFrame(hols, dtype=str)
holsa = hols.values.ravel()
#holsar = holsa.ravel() 
weekends = (5,6)
mtdy = "0" + str(mtod.month)
ytdy = str(date.today().year)
cwd = ytdy + "-" + mtdy


wdm2 = np.busday_offset(cwd,-2,roll='forward', holidays=list(holsa))
wdm1 = np.busday_offset(cwd,-1,roll='forward', holidays=list(holsa))
wd1 = np.busday_offset(cwd,0,roll='forward', holidays=list(holsa))
wd2 = np.busday_offset(cwd,1,roll='forward', holidays=list(holsa))
wd3 = np.busday_offset(cwd,2,roll='forward', holidays=list(holsa))
tdy = date.today()
#tdy = "2020-07-02"

if str(tdy) == str(wdm2):
    wd = "WD-2"
elif str(tdy) == str(wdm1):
    wd = "WD-1"
elif str(tdy) == str(wd1):
    wd = "WD1"
elif str(tdy) == str(wd2):
    wd = "WD2"
elif str(tdy) == str(wd3):
    wd = "WD3"
else:
    wd = tdy


# This function will Login to SAP from the SAP Logon window

coapath = r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx"
coa_df = pd.read_excel(coapath)
coa_df = pd.DataFrame(coa_df,columns= ['G/L Account'])
coa_df.to_clipboard(index=False)


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
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Vo2195pr@9112"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize
        
        if t_code == 3000 or t_code == "3000":
            session.findById("wnd[0]/tbar[0]/okcd").text = "faglb03"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/btn%_RACCT_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/usr/ctxtRBUKRS-LOW").text = "3000"
            session.findById("wnd[0]/usr/txtRYEAR").text = fisy
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)
            session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell(prd,"BALANCE")
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(2)
            coapath = r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx"
            coa_df = pd.read_excel(coapath)
            coa_df = pd.DataFrame(coa_df,columns= ['G/L Account'])
            coa_df.to_clipboard(index=False)
            
            session.findById("wnd[0]").sendVKey(33)
            session.findById("wnd[1]").sendVKey(71)
            time.sleep(1)
            
            session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/NRE_SEA_COE"
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
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3000_" + str(wd) + "_" + ttdy + ".xlsx"
            session.findById("wnd[1]").sendVKey(11)
            time.sleep(2)
            os.system("TASKKILL /F /IM saplogon.exe") 
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe")
            
        if t_code == 3505 or t_code == "3505":
            session.findById("wnd[0]/tbar[0]/okcd").text = "faglb03"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/btn%_RACCT_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/usr/ctxtRBUKRS-LOW").text = "3505"
            session.findById("wnd[0]/usr/txtRYEAR").text = fisy
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)
            session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell(prd,"BALANCE")
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(2)
            coapath = r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx"
            coa_df = pd.read_excel(coapath)
            coa_df = pd.DataFrame(coa_df,columns= ['G/L Account'])
            coa_df.to_clipboard(index=False)
            
            session.findById("wnd[0]").sendVKey(33)
            session.findById("wnd[1]").sendVKey(71)
            time.sleep(1)
            
            session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "/NRE_SEA_COE"
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
            session.findById("wnd[1]").sendVKey(0)
                              
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3505_" + str(wd) + "_" + ttdy + ".xlsx"
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


# In[3]:


saplogin(3505)


# In[5]:


import pandas as pd

from datetime import datetime, timedelta 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
tprd = datetime.today()+timedelta(days=65)
prd = tprd.strftime('%#m')
fprd = tprd.strftime('%m')
fis = datetime.now() + relativedelta(days=105)
fisy = fis.strftime('%Y')
sfisy = fis.strftime('%y')
sprd = (datetime.today()-relativedelta(days=10)).strftime('%b')
# print(fprd)
df1 = pd.read_excel(r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd + r"\3000_" + str(wd) + "_" + ttdy + ".xlsx", dtype={'Profit Center':str,'G/L Account':str})
# df2 = pd.read_excel(r"K:\FIN\Ivan\NRE\Data\3505_" + fisy + "_P" + fprd + "_WD-1 (2pm).xlsx", dtype={'Profit Center':str})

reg_df = pd.read_excel(r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx", sheet_name="Trad Part")
cha_df = pd.read_excel(r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx", sheet_name="Accounts", dtype={'G/L Account':str})

df2 = df1.merge(reg_df,how='left', on='Trading Part.BA')
df3 = df2.merge(cha_df,how='left', on='G/L Account')
pfc = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center")
pfc = pd.DataFrame(pfc,columns= ['ECC Profit Center','Business Unit.1','Level 1','Level 1 Description'])
pfc = pfc.drop_duplicates()
pfc = pfc.rename(columns={"ECC Profit Center":"Profit Center"})
df4 = df3.merge(pfc,how='left', on='Profit Center')
df4['LC Amount in USD'] = df4['Amount in local currency']/-1
df4['Exchange rate'] = -1
coa = pd.read_excel(r'K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx', sheet_name='BU')
coa2 = pd.read_excel(r'K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx', sheet_name='IDS',dtype=str)
dfc = df4.merge(coa, on="Business Unit.1",how="left")
dfc = dfc.merge(coa2, on="Level 1",how='left')
dfc['BU Description'] = np.where(dfc['BU_x'].isnull(),dfc['BU_y'],dfc['BU_x'])
dfc = dfc.drop(columns=['BU_x','BU_y'])
dfc = dfc.dropna(subset=['Company Code'])

dfc['Country'] = dfc['Country'].astype('category')
dfc['Country'] = dfc['Country'].cat.set_categories(['PHP','THA','SIN','MAL','VTN','OTA','MYA','ISA','PAK'])
dfc['Group Account Description'] = dfc['Group Account Description'].astype('category')
dfc['Group Account Description'] = dfc['Group Account Description'].cat.set_categories(['Gross Sales', 'Service Revenue', 'Distribution Margin', 'FX'])

pivdf = pd.pivot_table(dfc, index=['BU Description','Level 1 Description'],columns=['Group Account Description','Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False)

pivdf2 = pd.concat([
    d.append(d.sum().rename(("Total " + k,k)))
    for k,d in pivdf.groupby(['BU Description'])]).dropna(axis=0)


pivdf3 = pd.pivot_table(dfc2, index=['BU Description','Level 1 Description'],columns=['Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False).dropna(axis=0)
pivdf3 = pd.concat([
    d.append(d.sum().rename(("Total " + k ,k)))
    for k,d in pivdf3.groupby(['BU Description'])]).dropna(axis=0)
                    
wb = r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd + r"\3000_" + str(wd) + "_" + ttdy + ".xlsx"
with pd.ExcelWriter(wb) as ew:
    dfc.to_excel(ew, index=False, sheet_name="SAPfile") 
    pivdf2.to_excel(ew, sheet_name='Rev_Detail',engine='xlsxwriter')
    pivdf3.to_excel(ew, sheet_name='Net Sales',engine='xlsxwriter')
    workbook = ew.book
    worksheet = ew.sheets['Rev_Detail']
    worksheet2 = ew.sheets['Net Sales']
    num_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('A:B', 30, num_format)
    worksheet2.set_column('A:B', 30, num_format)
    worksheet.set_column('C:AZ', 12, num_format)
    worksheet2.set_column('C:AZ', 12, num_format)
    ew.save()


# In[6]:


tdy = datetime.today().strftime('%d_%m_%Y')
tprd = datetime.today()+timedelta(days=65)
prd = tprd.strftime('%#m')
fprd = tprd.strftime('%m')
fis = datetime.now() + relativedelta(days=105)
fisy = fis.strftime('%Y')
sfisy = fis.strftime('%y')
sprd = (datetime.today()-relativedelta(days=10)).strftime('%b')
# print(fprd)
# df1 = pd.read_excel(r"K:\FIN\Ivan\NRE\Data\3000_" + fisy + "_P" + fprd + "_WD-1 (2pm).xlsx", dtype={'Profit Center':str})
cdf = pd.read_excel(r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd + r"\3505_" + str(wd) + "_" + ttdy + ".xlsx", dtype={'Profit Center':str,'G/L Account':str})

reg_df = pd.read_excel(r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx", sheet_name="Trad Part")
cha_df = pd.read_excel(r"K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx", sheet_name="Accounts", dtype={'G/L Account':str})

cdf2 = cdf.merge(reg_df,how='left', on='Trading Part.BA')
cdf3 = cdf2.merge(cha_df,how='left', on='G/L Account')
pfc = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center")
pfc = pd.DataFrame(pfc,columns= ['ECC Profit Center','Business Unit.1','Level 1','Level 1 Description'])
pfc = pfc.drop_duplicates()
pfc = pfc.rename(columns={"ECC Profit Center":"Profit Center"})
cdf4 = cdf3.merge(pfc,how='left', on='Profit Center')
er = pd.read_excel(r"K:\FIN\Exchange rates\ExchRate_Summary " + sfisy + ".xlsx", sheet_name="Summary")
val = er.iloc[15,3+int(prd)]
                    
    
cdf4['LC Amount in USD'] = cdf4['Amount in local currency']/(-val)
cdf4['Exchange rate'] = (-val)
coa = pd.read_excel(r'K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx', sheet_name='BU')
coa2 = pd.read_excel(r'K:\FIN\Ivan\NRE\Mapping\Chart of account.xlsx', sheet_name='IDS', dtype=str)
df3c = cdf4.merge(coa, on="Business Unit.1",how="left")
df3c = df3c.merge(coa2, on="Level 1",how='left')
df3c['BU Description'] = np.where(df3c['BU_x'].isnull(),df3c['BU_y'],df3c['BU_x'])
df3c2 = df3c.drop(columns=['BU_x','BU_y'])
df3c2 = df3c2.dropna(subset=['Company Code'])

df3c2['Country'] = df3c2['Country'].astype('category')
df3c2['Country'] = df3c2['Country'].cat.set_categories(['PHP','THA','SIN','MAL','VTN','OTA','MYA','ISA','PAK'])
df3c2['Group Account Description'] = df3c2['Group Account Description'].astype('category')
df3c2['Group Account Description'] = df3c2['Group Account Description'].cat.set_categories(['Gross Sales', 'Service Revenue', 'Distribution Margin', 'FX'])

pivcdf = pd.pivot_table(df3c2, index=['BU Description','Level 1 Description'],columns=['Group Account Description','Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False)

pivcdf2 = pd.concat([
    d.append(d.sum().rename(("Total " + k,k)))
    for k,d in pivcdf.groupby(['BU Description'])]).dropna(axis=0)


pivcdf3 = pd.pivot_table(df3c2, index=['BU Description','Level 1 Description'],columns=['Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False).dropna(axis=0)
pivcdf3 = pd.concat([
    d.append(d.sum().rename(("Total " + k ,k)))
    for k,d in pivcdf3.groupby(['BU Description'])]).dropna(axis=0)

                    
eb = r"Y:\SEA Finance\NRE\Actual FY" + sfisy + "\\"  +  fprd + ". " + sprd + r"\3505_" + str(wd) + "_" + ttdy + ".xlsx"
with pd.ExcelWriter(eb) as ez:
    df3c2.to_excel(ez, index=False,sheet_name="SAPfile") 
    pivcdf2.to_excel(ez, sheet_name='Rev_Detail',engine='xlsxwriter')
    pivcdf3.to_excel(ez, sheet_name='Net Sales',engine='xlsxwriter')
    workbook = ez.book
    worksheet = ez.sheets['Rev_Detail']
    worksheet2 = ez.sheets['Net Sales']
    num_format = workbook.add_format({'num_format': '#,##0'})
    worksheet.set_column('A:B', 30, num_format)
    worksheet2.set_column('A:B', 30, num_format)
    worksheet.set_column('C:AZ', 12, num_format)
    worksheet2.set_column('C:AZ', 12, num_format)
    ez.save()


# In[ ]:


# # CTY = ['PHP','THA','SIN','MAL','VTN','OTA','MYA','ISA','PAK']
# from openpyxl.workbook import Workbook

# df2['Country'] = df2['Country'].astype('category')
# df2['Country'] = df2['Country'].cat.set_categories(['PHP','THA','SIN','MAL','VTN','OTA','MYA','ISA','PAK'])
# df2['Group Account Description'] = df2['Group Account Description'].astype('category')
# df2['Group Account Description'] = df2['Group Account Description'].cat.set_categories(['Gross Sales', 'Service Revenue', 'Distribution Margin', 'FX'])

# # pivdf = pd.pivot_table(df2, index=['BU Description','Level 1 Description'],columns=['Group Account','Group Account Description','Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False)
# pivdf = pd.pivot_table(df2, index=['BU Description','Level 1 Description'],columns=['Group Account Description','Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False)

# pivdf2 = pd.concat([
#     d.append(d.sum().rename(("Total " + k,k)))
#     for k,d in pivdf.groupby(['BU Description'])]).dropna(axis=0)


# pivdf3 = pd.pivot_table(df2, index=['BU Description','Level 1 Description'],columns=['Country'],values=['LC Amount in USD'], aggfunc=np.sum, fill_value=0, margins=True,margins_name='Total Business', dropna=False).dropna(axis=0)
# pivdf3 = pd.concat([
#     d.append(d.sum().rename(("Total " + k ,k)))
#     for k,d in pivdf3.groupby(['BU Description'])]).dropna(axis=0)


# wb = r"C:\Users\10299976\OneDrive - BD\Documents\My Received Files\NRE_Pivot.xlsx"
# with pd.ExcelWriter(wb) as ew:
#     pivdf2.to_excel(ew, sheet_name='Rev_Detail',engine='xlsxwriter')
#     pivdf3.to_excel(ew, sheet_name='Net Sales',engine='xlsxwriter')
#     workbook = ew.book
#     worksheet = ew.sheets['Rev_Detail']
#     worksheet2 = ew.sheets['Net Sales']
#     num_format = workbook.add_format({'num_format': '#,##0'})
#     worksheet.set_column('A:B', 30, num_format)
#     worksheet2.set_column('A:B', 30, num_format)
#     worksheet.set_column('C:AZ', 12, num_format)
#     worksheet2.set_column('C:AZ', 12, num_format)
#     ew.save()


# In[22]:


from datetime import datetime, timedelta 
from datetime import date

if date.today().day -14 >= 15 :
    tprd = date.today().month + 3
else:
    tprd = date.today().month + 2

prd = str(tprd).format('%#m')

prd


# In[ ]:




