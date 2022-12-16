import os
import pandas as pd
import dataframe_image as dfi
import PySimpleGUI as sg

def generate_report(month):
    
    subsplashReport = pd.read_csv(month + "/SS.csv")
    planningCenterReport = pd.read_csv(month + "/PCO.csv")

    subsplashReport = subsplashReport[['first_name', 'last_name', 'subfund', 'net_amount', 'frequency']]
    subsplashReport.rename(columns={'first_name':'First Name', 'last_name':'Last Name', 'subfund':'Sub Fund', 'net_amount':'Net Amount', 'frequency':'Frequency'}, inplace=True)

    planningCenterReport = planningCenterReport[['donor_first_name', 'donor_last_name', 'fund', 'net_amount']]
    planningCenterReport.rename(columns={'donor_first_name':"First Name", 'donor_last_name':"Last Name", 'fund':'Sub Fund', 'net_amount':'Net Amount'}, inplace=True)


    outputReport = pd.concat([planningCenterReport, subsplashReport])
    outputReport = outputReport.sort_values(by=['Sub Fund'])

    for subFund in outputReport['Sub Fund'].unique():
        individualReport = outputReport.loc[outputReport['Sub Fund'] == subFund]
        dfi.export(individualReport.style.hide(axis='index'), month + "/Individual Reports/" + subFund + " " + month + ".png")

        if(os.path.isfile('/Continuous Reports/'+subFund+' report')):
            #open excel file and add sheet to it
        else:
            print("idk")
            #create excel file and add the correct sheet
            

    outputReport.to_csv(month + "/" + month + "Staff Support Report.csv", index=False)

files = os.listdir(os.getcwd())
months = [f for f in files if os.path.isdir(os.getcwd()+'/'+f) and f != "environment"]

layout = [[sg.Text("Please Select A Month:")], [sg.Combo(months)],[sg.Button("Submit")]]


window = sg.Window("CityLight Report Generator", layout, margins=(300,300))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    if event == "Submit":
        generate_report(values[0])
        break

window.close()

#write to excel spreadsheet for each person and have each sheet be a month for that person and the main sheet be an overview page
