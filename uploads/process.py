import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from pivot_03 import rtf
import xlsxwriter

file01 = pd.read_excel('./uploads/details.xlsx')
file02 = pd.read_excel('./uploads/LCS tracker.xlsx')
file03 = pd.read_excel('./uploads/CAPT.xlsx')

print(os.getcwd())

def output(file01,file02,file03):
    print("xyz")
    df_new = file01
    dealers_01 = ['Others','Other','GENWORKS','SHREEJI HEALTHCARE','Genworks','OTHER']
    z = ['Cancelled']
    z1 = ['Cancelled','Pending']
    z2 = ['Pending','pending','PENDING']
    x1 = ['M3','m3']
    y1 = ['3/3/2017']
    flags_01 = df_new['Dealer Name'].isin(dealers_01) & df_new["Product: Tier 3 P&L"].notnull() & ~df_new['Order Status'].isin(z)  & df_new['Execution quarter'].notnull()
    table_01 = pd.pivot_table(df_new[flags_01],index=['Zone','Account'],columns=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)


    flags_02 = df_new["Product: Tier 3 P&L"].notnull() & df_new['Execution quarter'].notnull() & ~df_new['Order Status'].isin(z) & df_new['Order Status'].notnull()
    table_02 = pd.pivot_table(df_new[flags_02],index=['Zone','Account'],columns=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)

    xyz = ['GENWORKS','Genworks','SHREEJI HEALTHCARE']

    flags_03 =  df_new['Dealer Name'].isin(xyz) & df_new["Product: Tier 3 P&L"].notnull() & ~df_new['Order Status'].isin(z) & df_new['Execution quarter'].notnull() & df_new['Order Status'].notnull() 



    table_03 = pd.pivot_table(df_new[flags_03],index=['Zone'],columns=['Order Status'],values=['Net Equipment Order Value'],aggfunc=np.sum,margins=True)

    flags_04 = df_new["Product: Tier 3 P&L"].notnull() & df_new['Execution quarter'].notnull() & ~df_new['Order Status'].isin(z) & df_new['Order Status'].notnull()
    table_04 = pd.pivot_table(df_new[flags_04],index=['Zone'],columns=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)

    flags_05 = df_new["Product: Tier 3 P&L"].notnull() & df_new['Execution quarter'].notnull() & ~df_new['Order Status'].isin(z) & df_new['Order Status'].notnull()
    table_05 = pd.pivot_table(df_new[flags_05],index=['Product: Tier 3 P&L'],columns=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)


    flags_06 = df_new["Product: Tier 3 P&L"].notnull() & df_new['Execution quarter'].notnull() & ~df_new['Order Status'].isin(z) & df_new['Order Status'].notnull()
    table_06 = pd.pivot_table(df_new[flags_06],index=['Zone','Sales Rep'],columns=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)

    flags_07 = df_new["Product: Tier 3 P&L"].notnull() & df_new['Execution quarter'].notnull() & ~df_new['Order Status'].isin(z1) & df_new['Order Status'].notnull()
    table_07 = pd.pivot_table(df_new[flags_07],index=['Order Status'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)

    flags_08 =  df_new['Execution quarter'].notnull() & df_new['Order Status'].isin(z2) & df_new['Other Comments'].isin(x1)
    table_08 = pd.pivot_table(df_new[flags_08],index=['Zone','Account'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)

    flags_09 =  df_new['Ok to Bill Date'].isin(y1) 
    table_09 = pd.pivot_table(df_new[flags_09],index=['Account'],values=['Net Equipment Order Value'],
                           aggfunc=np.sum,margins=True)
    print(table_01)

    writer = pd.ExcelWriter('RTF_02.xlsx',engine='openpyxl')
    writer.table_01 = table_01
    writer.table_02 = table_02
    writer.table_03 = table_03
    writer.table_04 = table_04
    writer.table_05 = table_05
    writer.table_06 = table_06
    writer.table_07 = table_07
    writer.table_08 = table_08
    writer.table_09 = table_09
    table = rtf(file01,file02,file03)
    table_01.to_excel(writer,"Summary- By Region GW")
    table_02.to_excel(writer,"Summary- By Region")
    table_03.to_excel(writer,"Summary- GenWorks")
    table_04.to_excel(writer,"Summary_01")
    table_05.to_excel(writer,"Summary_02")
    table_06.to_excel(writer,"Summary_03")
    table_07.to_excel(writer,"Summary_04")
    table_08.to_excel(writer,"Summary_05")
    table_09.to_excel(writer,"Summary_06")
    table.to_excel(writer,"RTF all formulas")
    print("abc")
    writer.save()
    return

output(file01,file02,file03)