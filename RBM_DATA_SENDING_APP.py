
import pandas as pd
import numpy as np

import PySimpleGUI as sg
import time
import xlwings as xw
import glob
import pyautogui as pg
import shutil
sg.theme('DarkBlue')

layout = [
    [sg.T("Input Data Extraction File (From DB):", s=45,justification="l"), sg.I(key="-IN-", s=45), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
    [sg.Text("Input Year",s=45,justification="l"),sg.InputText( key='-DATE-')],
    [sg.Text("Choose Budget (will be in AGMT)",s=25,justification="l")],
    [sg.Listbox(values=['Act ', 'AP ', 'ASP_', 'May_LE ', 'Mar PMM ', 'Nov PMM ','Sep PMM_','SRC_'], size=(60, 10), select_mode='single', key='-BUDGET-')],
    [sg.T("Input Data Sending File:", s=45,justification="l"), sg.I(key="-SEND-", s=45), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
    [sg.T("Input Data Storage Folder:", s=45,justification="l"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],
    [sg.Text("Input Budget required from regions",s=45,justification="l"),sg.InputText( key='-FCST-')],
    [sg.Text("What are we putting in AGMT (agmt/none)",s=45,justification="l"),sg.InputText( key='-AGMT_DATA-')],
    [sg.Text("What are we putting in FCST (agmt/act/none)",s=45,justification="l"),sg.InputText( key='-FCST_DATA-')],
     [sg.Listbox(values=["East","Moscow","North West","Siberia","South","Ural","Volga"], size=(60, 10), select_mode='multiple', key='-REGIONS-')],
    [sg.Submit( )]
]

window = sg.Window('RBM SENDING APP', layout)

event, values = window.read()

extract_file=str(values['-IN-'])
year=str(values['-DATE-'])
budget=str(values['-BUDGET-'][0])
send_file=str(values['-SEND-'])
result_file=str(values['-FOLDER-'])
fcst=str(values['-FCST-'])
fcst_data=str(values['-FCST_DATA-']).lower()
agmt_data=str(values['-AGMT_DATA-']).lower()
regions_list=["East","Moscow","North West","Siberia","South","Ural","Volga"]
#regions_list=values['-REGIONS-']
window.close()

extract_file=extract_file.replace("\\","\\\\")
send_file=send_file.replace("\\","\\\\")
result_file=result_file.replace("\\","\\\\")

date="01.01."+year


chains_li=[]
penalty_li=[]
fcst_v_li=[]
fcst_b_li=[]
fcst_a_li=[]





wb = xw.Book(extract_file)
ws = wb.sheets["In"]
ws['D3'].value=date
ws['F3'].value="ALL"
ws['H3'].value=budget+"Volume"
ws['J3'].value=budget+"Budget"
ws['L3'].value=budget+"Activities"

wb.api.RefreshAll()
time.sleep(30)

ws = wb.sheets["Chain"]
chains=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value

east_chains=chains[chains['Region']=="East"]
mos_chains=chains[chains['Region']=="Moscow"]
nw_chains=chains[chains['Region']=="North West"]
sib_chains=chains[chains['Region']=="Siberia"]
south_chains=chains[chains['Region']=="South"]
ural_chains=chains[chains['Region']=="Ural"]
volga_chains=chains[chains['Region']=="Volga"]

ws = wb.sheets["Penalties"]
penalty=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value

east_penalty=penalty[penalty['Region']=="East"]
mos_penalty=penalty[penalty['Region']=="Moscow"]
nw_penalty=penalty[penalty['Region']=="North West"]
sib_penalty=penalty[penalty['Region']=="Siberia"]
south_penalty=penalty[penalty['Region']=="South"]
ural_penalty=penalty[penalty['Region']=="Ural"]
volga_penalty=penalty[penalty['Region']=="Volga"]

ws = wb.sheets["Act Volume"]
act_volume=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value

east_act_volume=act_volume[act_volume['Region']=="East"]
mos_act_volume=act_volume[act_volume['Region']=="Moscow"]
nw_act_volume=act_volume[act_volume['Region']=="North West"]
sib_act_volume=act_volume[act_volume['Region']=="Siberia"]
south_act_volume=act_volume[act_volume['Region']=="South"]
ural_act_volume=act_volume[act_volume['Region']=="Ural"]
volga_act_volume=act_volume[act_volume['Region']=="Volga"]

ws = wb.sheets["Act Budget"]
act_budget=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value
        
east_act_budget=act_budget[act_budget['Region']=="East"]
mos_act_budget=act_budget[act_budget['Region']=="Moscow"]
nw_act_budget=act_budget[act_budget['Region']=="North West"]
sib_act_budget=act_budget[act_budget['Region']=="Siberia"]
south_act_budget=act_budget[act_budget['Region']=="South"]
ural_act_budget=act_budget[act_budget['Region']=="Ural"]
volga_act_budget=act_budget[act_budget['Region']=="Volga"]

ws = wb.sheets["Act Activities"]
act_activities=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value
        
east_act_activities=act_activities[act_activities['Region']=="East"]
mos_act_activities=act_activities[act_activities['Region']=="Moscow"]
nw_act_activities=act_activities[act_activities['Region']=="North West"]
sib_act_activities=act_activities[act_activities['Region']=="Siberia"]
south_act_activities=act_activities[act_activities['Region']=="South"]
ural_act_activities=act_activities[act_activities['Region']=="Ural"]
volga_act_activities=act_activities[act_activities['Region']=="Volga"]


#----------------------------------------------------------------------
ws = wb.sheets["AGRM Volume"]
agmt_volume=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value

east_agmt_volume=agmt_volume[agmt_volume['Region']=="East"]
mos_agmt_volume=agmt_volume[agmt_volume['Region']=="Moscow"]
nw_agmt_volume=agmt_volume[agmt_volume['Region']=="North West"]
sib_agmt_volume=agmt_volume[agmt_volume['Region']=="Siberia"]
south_agmt_volume=agmt_volume[agmt_volume['Region']=="South"]
ural_agmt_volume=agmt_volume[agmt_volume['Region']=="Ural"]
volga_agmt_volume=agmt_volume[agmt_volume['Region']=="Volga"]

ws = wb.sheets["AGRM Budget"]
agmt_budget=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value
        
east_agmt_budget=agmt_budget[agmt_budget['Region']=="East"]
mos_agmt_budget=agmt_budget[agmt_budget['Region']=="Moscow"]
nw_agmt_budget=agmt_budget[agmt_budget['Region']=="North West"]
sib_agmt_budget=agmt_budget[agmt_budget['Region']=="Siberia"]
south_agmt_budget=agmt_budget[agmt_budget['Region']=="South"]
ural_agmt_budget=agmt_budget[agmt_budget['Region']=="Ural"]
volga_agmt_budget=agmt_budget[agmt_budget['Region']=="Volga"]

ws = wb.sheets["AGRM Activities"]
agmt_activities=ws.range('A1').options(pd.DataFrame, 
        header=1,
        index=False, 
        expand='table').value
        
east_agmt_activities=agmt_activities[agmt_activities['Region']=="East"]
mos_agmt_activities=agmt_activities[agmt_activities['Region']=="Moscow"]
nw_agmt_activities=agmt_activities[agmt_activities['Region']=="North West"]
sib_agmt_activities=agmt_activities[agmt_activities['Region']=="Siberia"]
south_agmt_activities=agmt_activities[agmt_activities['Region']=="South"]
ural_agmt_activities=agmt_activities[agmt_activities['Region']=="Ural"]
volga_agmt_activities=agmt_activities[agmt_activities['Region']=="Volga"]
wb.save()
wb.close()

for i in regions_list:
    print("Creating, ",i," File")
    source = send_file
    target = result_file + "\\RBM_data_collection_file_" + i+"_"+fcst+".xlsx"
    shutil.copy(source, target)
    wb = xw.Book(target)
    ws=wb.sheets["Manual"]
    ws['D9'].value=budget+"_"+year
    ws['D10'].value=budget+"_"+year
    ws['D11'].value=budget+"_"+year
    ws['D12'].value=fcst
    ws['D13'].value=fcst
    ws['D14'].value=fcst
    ws = wb.sheets["Chains"]
    
    if i=="East":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_chains
    elif i=="Moscow":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_chains
    if i=="North West":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_chains
    if i=="Siberia":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_chains
    if i=="South":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_chains
    if i=="Ural":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_chains
    if i=="Volga":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_chains
    
    ws = wb.sheets["Penalty"]
    
    if i=="East":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_penalty
    elif i=="Moscow":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_penalty
    if i=="North West":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_penalty
    if i=="Siberia":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_penalty
    if i=="South":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_penalty
    if i=="Ural":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_penalty
    if i=="Volga":
        ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_penalty

    if agmt_data=="agmt":
        ws = wb.sheets["AGMT LE Volume"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_volume
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_volume
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_volume
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_volume
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_volume
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_volume
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_volume

        ws = wb.sheets["AGMT LE Budget"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_budget
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_budget
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_budget
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_budget
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_budget
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_budget
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_budget   
        
        ws = wb.sheets["AGMT LE Activities"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_activities
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_activities
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_activities
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_activities
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_activities
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_activities
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_activities
    if agmt_data=="none":
        pass





    if fcst_data=="act":
        ws = wb.sheets["Fcst LE Volume"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_act_volume
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_act_volume
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_act_volume
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_act_volume
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_act_volume
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_act_volume
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_act_volume

        ws = wb.sheets["Fcst LE Budget"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_act_budget
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_act_budget
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_act_budget
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_act_budget
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_act_budget
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_act_budget
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_act_budget   
        
        ws = wb.sheets["Fcst LE Activities"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_act_activities
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_act_activities
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_act_activities
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_act_activities
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_act_activities
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_act_activities
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_act_activities
    if fcst_data=="agmt":
        ws = wb.sheets["Fcst LE Volume"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_volume
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_volume
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_volume
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_volume
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_volume
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_volume
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_volume

        ws = wb.sheets["Fcst LE Budget"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_budget
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_budget
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_budget
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_budget
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_budget
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_budget
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_budget   
        
        ws = wb.sheets["Fcst LE Activities"]
        
        if i=="East":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = east_agmt_activities
        elif i=="Moscow":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = mos_agmt_activities
        if i=="North West":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = nw_agmt_activities
        if i=="Siberia":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = sib_agmt_activities
        if i=="South":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = south_agmt_activities
        if i=="Ural":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = ural_agmt_activities
        if i=="Volga":
            ws["A3"].options(pd.DataFrame, header=0, index=False, expand='table').value = volga_agmt_activities
    if fcst_data=="none":
        pass







    wb.api.RefreshAll()
    time.sleep(60)
    
    wb.api.RefreshAll()
    time.sleep(40)
    wb.save()
    wb.close()

   
        


