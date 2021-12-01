#!/usr/bin/env python
# coding: utf-8

# In[ ]:


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
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Vo2195pr@9112"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").maximize
        
        if t_code == "ke30" or t_code == "KE30":
            session.findById("wnd[0]/tbar[0]/okcd").text = "ke30"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "bd01"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[0]/shellcont/shell").doubleClickNode("000000001001")
            session.findById("wnd[0]/usr/ctxtPAR_04").text = "10"
            session.findById("wnd[0]/usr/ctxtPAR_02").text = fisy
            session.findById("wnd[0]/usr/ctxtPAR_05-LOW").text = "1"
            session.findById("wnd[0]/usr/ctxtPAR_05-HIGH").text = "12"
            session.findById("wnd[0]/usr/ctxtPAR_01-LOW").text = "3000"
            session.findById("wnd[0]/usr/ctxtPAR_08-LOW").text = "100"
            session.findById("wnd[0]/usr/ctxtPAR_08-HIGH").text = "999"
            
            session.findById("wnd[0]/usr/btn%_PAR_03_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1024"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "3091"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "3092"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1147"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "1149"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "3350"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "sg01"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "1014"
            session.findById("wnd[1]").sendVKey(82)
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "2331"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1150"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "1008"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "th01"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "z654"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "z656"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "my01"
            session.findById("wnd[1]").sendVKey(82)
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ph01"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "vn01"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "id01"
            session.findById("wnd[1]").sendVKey(8)
            
            
            session.findById("wnd[0]/usr/radALV1").select()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            session.findById("wnd[0]/tbar[1]/btn[33]").press()
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
            
#             session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "LCOL019T002"
#             session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("LCOL005K001")
#             session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
#             session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&FILTER")
#             session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "pk"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "sg"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "th"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "ph"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "my"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "vn"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "la"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "bn"
#             session.findById("wnd[2]").sendVKey(82)
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "mm"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "id"
#             session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "kh"

#             session.findById("wnd[2]/tbar[0]/btn[8]").press()
#             session.findById("wnd[1]/tbar[0]/btn[0]").press()
                       
            session.findById("wnd[0]").sendVKey(43)           
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\KE30\Data"
            
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "KE30_3000_" + fisy + ".xlsx"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe")
            
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nke30"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/shellcont/shell").doubleClickNode("000000001001")
            session.findById("wnd[0]/usr/ctxtPAR_04").text = "B0"
            session.findById("wnd[0]/usr/ctxtPAR_02").text = fisy
            session.findById("wnd[0]/usr/ctxtPAR_05-LOW").text = "1"
            session.findById("wnd[0]/usr/ctxtPAR_05-HIGH").text = "12"
            session.findById("wnd[0]/usr/ctxtPAR_01-LOW").text = "3505"
            session.findById("wnd[0]/usr/ctxtPAR_08-LOW").text = "100"
            session.findById("wnd[0]/usr/ctxtPAR_08-HIGH").text = "999"
#             session.findById("wnd[0]/usr/radALV1").select()
            session.findById("wnd[0]/usr/btn%_PAR_03_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]").sendVKey(16)
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]").sendVKey(8)
            session.findById("wnd[0]/tbar[1]/btn[33]").press()
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
            
            session.findById("wnd[0]").sendVKey(43)           
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\KE30\Data"
            
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "KE30_3505_" + fisy + ".xlsx"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe")
            
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/usr/ctxtPAR_04").text = "10"
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[33]").press()
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
            
            session.findById("wnd[0]").sendVKey(43)           
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\KE30\Data"
            
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "KE30_3505_IDR_" + fisy + ".xlsx"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()          
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


# In[ ]:


saplogin("ke30")


# In[1]:


import pandas as pd
from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=95)
fisy = fis.strftime('%Y')

pth3000 = r"K:\FIN\Ivan\KE30\Data\\" + "KE30_3000_" + fisy + ".xlsx"
pth3505 = r"K:\FIN\Ivan\KE30\Data\\" + "KE30_3505_" + fisy + ".xlsx"

main_df = pd.read_excel(pth3000, dtype=str)
isa_df = pd.read_excel(pth3505, dtype=str)

df1 = main_df.append(isa_df)
df1['Profit Center'] = df1['Profit Center'].str[5:]


filter_list = ['BN','LA','MM','MY','PH','PK','TH', 'ID','SG','KH','VN']
df1 = df1[df1.Country.isin(filter_list)]
# df1 = df1.astype(str)

pfc = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center", dtype=str)
pfc1 = pd.DataFrame(pfc,columns= ['ECC Profit Center','Business Unit.1'])
pfc1 = pfc1.rename(columns={'ECC Profit Center': 'Profit Center','Business Unit.1':'WWD'})

pfc1 = pfc1.astype(str)
df2 = df1.merge(pfc1,how='left', on="Profit Center")


wb = "K:\FIN\Ivan\KE30\Data\Total_KE30_" + fisy + ".xlsx"
with pd.ExcelWriter(wb) as ew:
    df2.to_excel(ew, index=False) 


# In[4]:


import pandas as pdw
from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=95)
fisy = fis.strftime('%Y')

pth3505 = r"K:\FIN\Ivan\KE30\Data\\" + "KE30_3505_IDR_" + fisy + ".xlsx"
isa_df2 = pd.read_excel(pth3505, dtype=str)
isa_df2['Profit Center'] = isa_df2['Profit Center'].str[5:]

pfc = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center", dtype=str)
pfc = pd.DataFrame(pfc,columns= ['ECC Profit Center','Business Unit.1'])
pfc = pfc.rename(columns={'ECC Profit Center': 'Profit Center','Business Unit.1':'WWD'})
df3 = isa_df2.merge(pfc,how='left', on="Profit Center")

with pd.ExcelWriter(pth3505) as ew:
    df3.to_excel(ew, index=False) 


# # In[5]:


# # import numpy as np
# # import calendar
# # from datetime import datetime 
# # from datetime import date
# # from dateutil.relativedelta import relativedelta

# # now = datetime.now()
# # tdy = int(date.today().strftime('%d'))
# # nday= int(str(calendar.monthrange(now.year,now.month))[4:6])/2

# # if tdy >= nday:
    # # fis = datetime.now() + relativedelta(days=85)
    # # fisd = fis.strftime('0%m.%Y')
# # else:
    # # fis = datetime.now() + relativedelta(days=61)
    # # fisd = fis.strftime('0%m.%Y')
    
# # ds = (df3['Period/year'] == fisd)
# # ds
# # dt = df3[ds]
# # dt['Net Revenue Period'] = pd.to_numeric(dt['Net Revenue Period'],errors='coerce')
# # dt = dt[~dt.Country.isna()]

# # piv3505 = pd.pivot_table(dt,index=["WWD","Segment.1"],values=['Net Revenue Period'], columns=['Country.1'], aggfunc=np.sum, fill_value=0, margins=True, margins_name='~ Grand Total ~')
# # piv3505 = pd.concat([
    # # d.append(d.sum().rename((k + 'Subtotal', ''))) 
    # # for k, d in piv3505.groupby('WWD')])


# # In[6]:


# # import numpy as np

# # now = datetime.now()
# # tdy = int(date.today().strftime('%d'))
# # nday= int(str(calendar.monthrange(now.year,now.month))[4:6])/2

# # if tdy >= nday:
    # # fis = datetime.now() + relativedelta(days=85)
    # # fisd = fis.strftime('0%m.%Y')
# # else:
    # # fis = datetime.now() + relativedelta(days=61)
    # # fisd = fis.strftime('0%m.%Y')

    
# # ds = (df3['Period/year'] == fisd)

# # dtgp = df3[ds]
# # dtgp['Std Gross Profit Period'] = pd.to_numeric(dtgp['Std Gross Profit Period'],errors='coerce')
# # dtgp = dtgp[~dtgp.Country.isna()]

# # piv3505gp = pd.pivot_table(dtgp,index=["WWD","Segment.1"],values=['Std Gross Profit Period'], columns=['Country.1'], aggfunc=np.sum, fill_value=0, margins=True, margins_name='~ Grand Total ~')
# # piv3505gp = pd.concat([
    # # d.append(d.sum().rename((k + 'Subtotal', ''))) 
    # # for k, d in piv3505gp.groupby('WWD')])


# # In[7]:


# # import numpy as np

# # now = datetime.now()
# # tdy = int(date.today().strftime('%d'))
# # nday= int(str(calendar.monthrange(now.year,now.month))[4:6])/2

# # if tdy >= nday:
    # # fis = datetime.now() + relativedelta(days=85)
    # # fisd = fis.strftime('0%m.%Y')
# # else:
    # # fis = datetime.now() + relativedelta(days=61)
    # # fisd = fis.strftime('0%m.%Y')

    
# # ds1 = (df2['Period/year'] == fisd)

# # dt1 = df2[ds1]
# # dt1['Net Revenue Period'] = pd.to_numeric(dt1['Net Revenue Period'],errors='coerce')

# # dt1 = dt1[~dt1.Country.isna()]

# # piv = pd.pivot_table(dt1,index=["WWD","Segment.1"],values=['Net Revenue Period'], columns=['Country.1'], aggfunc=np.sum, fill_value=0, margins=True, margins_name= '~ Grand Total ~')
# # piv = pd.concat([
    # # d.append(d.sum().rename((k + ' SubTotal',''))) 
    # # for k, d in piv.groupby('WWD')])


# # In[8]:


# # import numpy as np

# # now = datetime.now()
# # tdy = int(date.today().strftime('%d'))
# # nday= int(str(calendar.monthrange(now.year,now.month))[4:6])/2

# # if tdy >= nday:
    # # fis = datetime.now() + relativedelta(days=85)
    # # fisd = fis.strftime('0%m.%Y')
# # else:
    # # fis = datetime.now() + relativedelta(days=61)
    # # fisd = fis.strftime('0%m.%Y')

    
# # ds1 = (df2['Period/year'] == fisd)

# # dt1gp = df2[ds1]
# # dt1gp['Std Gross Profit Period'] = pd.to_numeric(dt1gp['Std Gross Profit Period'],errors='coerce')

# # dt1gp = dt1gp[~dt1gp.Country.isna()]

# # pivgp = pd.pivot_table(dt1gp,index=["WWD","Segment.1"],values=['Std Gross Profit Period'], columns=['Country.1'], aggfunc=np.sum, fill_value=0, margins=True, margins_name='~ Grand Total ~')
# # pivgp = pd.concat([
    # # d.append(d.sum().rename((k + ' SubTotal',''))) 
    # # for k, d in pivgp.groupby('WWD')])


# # # In[9]:


# # tdy = datetime.today().strftime('%d_%m_%Y')

# # wb = "K:\FIN\Ivan\KE30\KE30_MTD_" + tdy + ".xlsx"

# # with pd.ExcelWriter(wb) as ew:
    # # piv.to_excel(ew, sheet_name='All_USD_BU_REV', engine='xlsxwriter')
    # # pivgp.to_excel(ew, sheet_name='ALL_USD_BU_GP', engine='xlsxwriter')
    # # piv3505.to_excel(ew, sheet_name='3505_IDR_BU_REV',engine='xlsxwriter')
    # # piv3505gp.to_excel(ew, sheet_name='3505_IDR_BU_GP', engine='xlsxwriter')


# # In[11]:
import win32com.client
import pandas as pd
from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=95)
fisy = fis.strftime('%Y')

xl=win32com.client.Dispatch('Excel.Application')
xl.Workbooks.Open(Filename=r'K:\FIN\Ivan\KE30\Data\KE30_MTD.xlsb', ReadOnly=1)
xl.Application.Run('ThisWorkbook.MTD_report')
xl.Application.Quit()
del xl


import win32com.client
import pandas as pd
from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

tdy = datetime.today().strftime('%d_%m_%Y')
fis = datetime.now() + relativedelta(days=95)
fisy = fis.strftime('%Y')

xl2=win32com.client.Dispatch('Excel.Application')
xl2.Workbooks.Open(Filename=r'K:\FIN\Ivan\KE30\KE30_' + fisy + '_Pivot.xlsb', ReadOnly=1)
xl2.Application.Run('ThisWorkbook.refresh')
xl2.Application.Quit()
del xl2


# In[ ]:




# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




