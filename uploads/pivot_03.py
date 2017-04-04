import pandas as pd
import numpy as np

def rtf(file01,file02,file03):
    df = file01
    table = pd.pivot_table(df[df['Order Status']=='Pending'],index=['Zone',u'Siebel Quote / OEP Number','Account'],values=['Net Equipment Order Value'],aggfunc=np.sum)

    #Read additional files and set their index for merge
    lcs = file02


    tcols_lcs = ['LCS OEP Status','Site Readiness','Customer Readiness','Customer Instruction Remarks','Advance comments','PD Comments','REASON WHY CHEQUE IS NOT RECEIPTED','Sales person name','Commercial officers','Incomplete KYC','Existing customer?','CRD','Requested On-Site Date','Backordered status','Genworks','Location','Billing Status']
    #tcols_capt = ['Payment Terms','CAPT','CAP T#']
    cols_lcs = [u"REQUEST #",'OEP Status','SITE READINESS','CUSTOMER INSTRUCTION FOR SHIPMENT RECEIVED','CUSTOMER INSTRUCTION COMMENT','ADVANCE COMMENTS','PD COMMENTS','REASON WHY CHEQUE IS NOT RECEIPTED','LOGGED BY','ADVANCE STEP APPROVERS','COMMENTS-EXECUTION TEAM','EXISTING CUSTOMER']
    cols_df = [u"Siebel Quote / OEP Number",'CRD','Requested On-Site Date','Equipment Status','Dealer Name','Delivery Location','Order Remarks by PJM']
    cols_df1 = ['OEP Status','SITE READINESS','CUSTOMER INSTRUCTION FOR SHIPMENT RECEIVED','CUSTOMER INSTRUCTION COMMENT','ADVANCE COMMENTS','PD COMMENTS','REASON WHY CHEQUE IS NOT RECEIPTED','LOGGED BY','ADVANCE STEP APPROVERS','COMMENTS-EXECUTION TEAM','EXISTING CUSTOMER','CRD','Requested On-Site Date','Equipment Status','Dealer Name','Delivery Location','Order Remarks by PJM']


    cols_final = ['Zone',u'Siebel Quote / OEP Number','Account','Net Equipment Order Value','LCS OEP Status','CRD','Requested On-Site Date','Site Readiness','Customer Readiness','Customer Instruction Remarks','Advance comments','PD Comments','REASON WHY CHEQUE IS NOT RECEIPTED','Backordered status','Sales person name','Commercial officers','Incomplete KYC','Genworks','Location','Existing customer?','Billing Status']



    df = df[(df['Order Status']=='Pending') & (df['Execution quarter'] == "Q1'17")]
    df = df[cols_df]
    df.drop_duplicates(cols_df,keep = 'last',inplace=True)
    df = df.set_index(u"Siebel Quote / OEP Number")
    lcs = lcs[cols_lcs]
    lcs = lcs.reset_index()
    lcs = lcs.set_index(u"REQUEST #")
    lcs = lcs.merge(df,left_index=True,right_index=True)
    lcs.drop(lcs.columns[0],axis=1,inplace=True)
    lcs.index.name = 'RNo'
    lcs = lcs.reset_index()
    lcs.drop_duplicates(subset='RNo',keep='first', inplace=True)
    lcs = lcs.set_index('RNo')
    lcs.to_excel('lcs.xlsx','Sheet1')


    table = table.reset_index()
    table = table.set_index(u"Siebel Quote / OEP Number")


    #table = table.merge(df,left_index=True,right_index=True)
    table[tcols_lcs] = lcs[cols_df1]
    #table[tcols_capt] = capt[cols_capt]



    table = table.reset_index()
    table = table[cols_final]
    table = table.set_index(['Zone',u'Siebel Quote / OEP Number','Account'])
    return table
