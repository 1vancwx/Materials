#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import win32com.client
import sys
import subprocess
import time
import os
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

from datetime import datetime 
from datetime import date
from dateutil.relativedelta import relativedelta

fis = datetime.now() + relativedelta(days=105)
fisy = fis.strftime('%Y')
prd = datetime.now() + relativedelta(days=60)
fism = prd.strftime('%m')

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
        
        
        if t_code == "asset" or t_code == "ASSET":
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87011964"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "3000"
            session.findById("wnd[0]/usr/ctxtBERDATUM").text = "30.09." + fisy
            session.findById("wnd[0]/usr/radXEINZEL").select()
            session.findById("wnd[0]/usr/ctxtSRTVR").text = "Z998"
            session.findById("wnd[0]").sendVKey(8)
            session.findById("wnd[0]").sendVKey(32)
            session.findById("wnd[1]").sendVKey(7)
            session.findById("wnd[1]").sendVKey(7)
            session.findById("wnd[1]").sendVKey(7)
            session.findById("wnd[1]").sendVKey(0)			
            session.findById("wnd[0]").sendVKey(9)
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "K:\FIN\Ivan\Depreciation\Data"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SAP_Total_Asset.txt"
            
            session.findById("wnd[1]").sendVKey(11)
            time.sleep(5)
            os.system("TASKKILL /F /IM EXCEL.exe")
            
            cf1 = pd.read_csv('K:\FIN\Ivan\Depreciation\Data\SAP_Total_Asset.txt', skiprows=8, error_bad_lines=False, sep='\t', dtype=str,encoding='cp1252') 
            cf1 = pd.DataFrame(cf1, columns=['Asset','Asset description','Cap.date','    Acquis.val.','     Accum.dep.','      Book val.','BusA','Profit Ctr'])
            
            a = ['3092','3348','3976', 'Z654','Z656'] 
            cf1 = cf1[~cf1['BusA'].isin(a)]
            #cf1 = cf1[~cf1['      Book val.'].str.contains("0.00")]
            cf1 = cf1.dropna(subset=['BusA'])
            wf = "K:\FIN\Ivan\Depreciation\Data\Total_Asset.xlsx"
            with pd.ExcelWriter(wf) as eq:
                cf1.to_excel(eq, index=False) 
            
            # aw01n  
            df1 = pd.read_excel(r"K:\FIN\Ivan\Depreciation\Data\Total_Asset.xlsx")
            df1 = pd.DataFrame(df1,columns= ['Asset', 'Profit Ctr', 'BusA'])
            ddf1 = df1.rename(columns={'Profit Ctr': 'Profit Center','BusA':'Business Area'})
            session.findById("wnd[0]/tbar[0]/okcd").text = "/naw01n"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/subIDC_SELECT1:AW01N:0201/ctxtANLA-BUKRS").text = "3000"
            
            adf = []
            for i in ddf1["Asset"]:
                
                session.findById("wnd[0]/usr/subIDC_SELECT1:AW01N:0201/ctxtANLA-ANLN1").text = i
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE").select()
                time.sleep(1)
                    
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("ICONTEXT")
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("PERAF")
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("NAFAZ")
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("AAFAZ")                    
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("MAFAZ")
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("AUFWZ")
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectColumn("WAERS")

                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").contextmenu()
                session.findById("wnd[0]/usr/tabsIDC_TABSTRIP/tabpIDC_BWERTE/ssubIDC_SUBSCREEN:AW01N:0302/cntlIDC_GRID_PLAN/shellcont/shell/shellcont[1]/shell").selectContextMenuItemByText('Copy Text')
                time.sleep(1)
                columns = ["Status", "Period", "Ord. Dep.", "Unplanned Dep.", "Reserve", "Revaluate", "Curr"]
                df2 = pd.read_clipboard(header=None, names=columns)           
                df2 = df2.dropna(subset=['Status'])
                df2["Asset"] = i
                

                adf.append(df2)
                time.sleep(1)
            adf2 = pd.concat(adf)
			
            
        adf2 = adf2.merge(ddf1,how='left', on="Asset")
        pc1 = pd.read_excel(r"K:\FIN\Ivan\MRA\SEA Reltio hierarachy products.xlsx", sheet_name="Profit Center",dtype=str)
        pc1 = pd.DataFrame(pc1,columns= ['ECC Profit Center', 'Product Description', 'Level 3', 'Level 3 Description', 'Business Unit.1'])
        pc2 = pc1.rename(columns={'ECC Profit Center':'Profit Center','Business Unit.1':'BU'})
        pc2 = pc2.astype(str)
        adf2 = adf2.astype(str)
        dpf = adf2.merge(pc2,how='left', on="Profit Center")

        dpf['Ord. Dep.'] = dpf['Ord. Dep.'].str[:-1]
        dpf = dpf.rename(columns={'Level 3':'Level3','Level 3 Description':'Level3 Description'})
        dpf = dpf.drop_duplicates()

        wb = "K:\FIN\Ivan\Depreciation\Data\YTD_Depreciation_" + fisy + ".xlsx"
        with pd.ExcelWriter(wb) as ew:
            dpf.to_excel(ew, index=False)  

            
        time.sleep(2)
        os.system("TASKKILL /F /IM saplogon.exe") 
        
    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


# In[ ]:


saplogin("asset")


# In[1]:


import win32com.client

xl=win32com.client.Dispatch('Excel.Application')
xl.Workbooks.Open(Filename=r'K:\FIN\Ivan\Depreciation\YTD_Depreciation_Pivot.xlsb', ReadOnly=1)
xl.Application.Run('ThisWorkbook.refresh_source')
xl.Application.Quit()
del xl


# In[ ]:




