#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 02 12:53:35 2014

@author: Wiebke.Toussaint
"""

import pandas as pd
from pandas import ExcelWriter
from datetime import date
from openpyxl.reader.excel import load_workbook
import MyFunx, AllData

today = date.today()

lastmonth = today.month - 3
nextmonth = today.month + 1

Visibility = AllData.InboundData(lastmonth, nextmonth, today)

Samples = pd.ExcelFile(u'05_Samples\\SampleTrack.xlsx').parse(u'Master', skiprows = 0, index = None, encoding='utf-8')
Samples = Samples[[u'SKU',u'03_SampleRoom_TO_studio']]
Samples.loc[Samples[u'03_SampleRoom_TO_studio'].notnull(),u'SampleCount'] = 1

Vis = pd.merge(Visibility, Samples, on = u'SKU', how = u'left', sort = False )

V = Vis.sort(['Date received','Date booked','POs'], inplace = False, na_position = 'first')
V = V[['GLYear','GLMonth','GLDay','Buyer', 'UnitCost','TotalUnits','TotalCost','SKU','SimpleName','ProcurementStatus','Category','Supplier','DeliveryDue','POs','BP Qty','Ref','Status','Date booked','Partial delivery','Date received','LastQCed','Qty Counted','Qty Damaged','SampleCount','Oversupply','OTDLastReceived','Qty Received','Qty PutAway','ActualGoLiveDate']]
 
today = date.today()

#==============================================================================
# Generate MerchTrack Output Data
#==============================================================================
def format():
    worksheet.set_column('A:C', 8 )
    worksheet.set_column('D:D', 12 )
    worksheet.set_column('E:G', 10 )
    worksheet.set_column('H:H', 16 )
    worksheet.set_column('I:I', 32 )
    worksheet.set_column('J:M', 20 )
    worksheet.set_column('N:O', 8 )
    worksheet.set_column('P:U', 18 )
    worksheet.set_column('V:X', 12 )
    worksheet.set_column('Y:Y', 18 )
    worksheet.set_column('Z:AA', 12 )
    worksheet.set_column('AB:AB', 18 )

writer1 = ExcelWriter('04_Visibility\\Visibility ' + str(today) + '.xlsx')
V.to_excel(writer1, 'MASTER', index = False, encoding = 'utf-8' )   
workbook = writer1.book
worksheet = writer1.sheets['MASTER']
format()
Buyers = pd.Series(V[u'Buyer'].unique())
for b in Buyers:
    DataName = V[V['Buyer']==b]
    DataName.to_excel(writer1, b , index = False )   
    workbook = writer1.book
    worksheet = writer1.sheets[b]
    format()
writer1.save()

doc_name = u'Visibility Report '
part = u'04_Visibility\\Visibility ' + str(today) + '.xlsx'
message = u'Visibility Report' + str(today)
maillist = "MailList_Merch.txt"

MyFunx.send_message(doc_name, message, part, maillist)

#==============================================================================
# Generate WHTrack Output Data
#==============================================================================
WHTrack = V[['SKU','SimpleName','Category','Supplier','POs','DeliveryDue','Date booked','Date received','LastQCed',\
'OTDLastReceived','UnitCost','TotalUnits','TotalCost','BP Qty','Qty Counted','Qty Damaged','Qty Received','Qty PutAway',\
'Partial delivery','Buyer','ProcurementStatus','Status','Ref','GLMonth']]
WHTrack = WHTrack.sort(columns = ['Date received','POs','LastQCed'], ascending = True, na_position = 'last', inplace = False)
WHTrack.name = 'WHTrack'

Backlog = WHTrack.dropna(subset = ['Date received'], inplace = False)
Backlog = Backlog[(Backlog['LastQCed'].isnull()==True) & (Backlog['OTDLastReceived'].isnull()==True)]
Backlog.name = 'Backlog'

def format1():
    worksheet.set_column('A:A', 15 )
    worksheet.set_column('B:B', 40 )
    worksheet.set_column('D:C', 15 )
    worksheet.set_column('E:E', 10 )
    worksheet.set_column('F:J', 18 )
    worksheet.set_column('K:R', 10 )
    worksheet.set_column('S:W', 20 )
    worksheet.set_column('X:X', 6 )

writer2 = ExcelWriter('04_Visibility\\WHTrack ' + str(today) + '.xlsx')
Track = [WHTrack,Backlog]
for t in Track:
    t.to_excel(writer2, t.name , index = False )   
    workbook = writer2.book
    worksheet = writer2.sheets[t.name]
    format1()  
writer2.save()
   
doc_name = 'WH Stock Track Report '
part = '04_Visibility\\WHTrack ' + str(today) + '.xlsx'
message = 'Spree Stock Tracking on ' + str(today)
maillist = "MailList_WH.txt"

MyFunx.send_message(doc_name, message, part, maillist)

#==============================================================================
# Generate ProductTrack QuickStats
#==============================================================================

V = Vis[(Vis['Ref'].str.contains(u"sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud|not OK")==False) | (Vis['Ref'].isnull()==True)] 

#ProductTrack
SimplesCount = V.groupby('GLMonth').agg({'TotalUnits':'count','POs': 'count',\
'Date booked':'count', 'Date received':'count','Qty Counted':'count', 'Qty Received':'count','Qty PutAway':'count'})
SimplesCount.sort_index(ascending = True, inplace = True)
SimplesCount = SimplesCount[['TotalUnits','POs', 'Date booked', 'Date received','Qty Counted','Qty Received','Qty PutAway']]
SimplesCount.rename(columns={'TotalUnits':'Simples Planned','POs':'Simples on BP','Date booked':'Simples booked',\
'Date received':'Simples received UNCHECKED','Qty Counted':'Simples QCed','Qty Received':'Simples taken in by OTD','Qty PutAway':'Simples in OTD WH'}, inplace = True)
SimplesCount.name = "Simples Count"

UnitsCount = V.groupby('GLMonth').agg({'TotalUnits':'sum','Qty Counted':'sum','SampleCount':'sum','Qty Damaged':'sum','Qty Received':'sum','Qty PutAway':'sum'})
UnitsCount.sort_index(ascending = True, inplace = True)
UnitsCount = UnitsCount[['TotalUnits','Qty Counted','SampleCount','Qty Damaged','Qty Received','Qty PutAway']]
UnitsCount.rename(columns={'TotalUnits':'Units Planned','Qty Counted':'Units QCed','Qty Received':'Units taken in by OTD','Qty PutAway':'Units in OTD WH'}, inplace = True)
UnitsCount.name = "Units Count"

POCount = V.dropna(subset = ['POs']).sort(columns = ['LastQCed','OTDLastReceived'], na_position = 'last')
POCount = POCount.drop_duplicates(subset = ['POs','LastQCed'])
POCount['duplicate'] = POCount.duplicated(subset = 'POs')
POCount.loc[POCount['LastQCed'].notnull(), 'atWH'] = 1
POCount = POCount.groupby('GLMonth')

POCount = POCount.apply(lambda x : pd.Series(dict(
POTotal = len(x.loc[x.duplicate==False]), 
POBooked = len(x.loc[(x.duplicate==False) & x['Date booked'].notnull()]), 
POReceived = len(x.loc[(x.duplicate==False) & x['Date received'].notnull()]),
PONotRec = len(x.loc[(x.duplicate==False) & x['Date received'].isnull()]),
Partial = len(x.loc[(x.duplicate==True) & x['Date received'].isnull()]), 
LastQCed = len(x.loc[(x.duplicate==False) & x['LastQCed'].notnull()]), 
OTDLastReceived = len(x.loc[(x.duplicate==False) & x['OTDLastReceived'].notnull()]))))

POCount = POCount[['POTotal', 'POBooked', 'POReceived','PONotRec','Partial','LastQCed','OTDLastReceived']]
POCount.rename(columns={'POTotal':'POs on BP','POBooked':'POs booked','POReceived':'POs received UNCHECKED','Partial':'Partial delivery','LastQCed':'POs QCed','OTDLastReceived':'POs in WH'}, inplace = True)
POCount = (POCount.T/list(POCount['POs on BP'])).T
POCount.name = "PO Count"

#SKUs not processed
V['TotalCost'] = V['UnitCost']*V['TotalUnits']
V['NoBP'] = V['Status'].isnull() #SKUs not on Brightpearl
V['NBND'] = V['Date received'].isnull() & V['Date booked'].isnull() #SKUs not booked not delivered
V['ND'] = V['Date received'].isnull() #SKUs not delivered
V['Partial'] = V['LastQCed'].isnull() #SKUs not QCed
V['NOTD'] = V['OTDLastReceived'].isnull() #SKUs not received by OTD

#Cost of SKUs not processed / WorkingCapital
NoBP = V[V['NoBP']==True].groupby(V['GLMonth']).sum()['TotalCost'] #Cost of SKUs not on BP / month
NBND = V[V['NBND']==True].groupby(V['GLMonth']).sum()['TotalCost'] - NoBP #Cost of SKUs not booked not delivered
BND = V[V['ND']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND) #Cost of SKUs not delivered / month
Partial = V[V['Partial']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND + BND)  #Cost of SKUs not QCed
#NOTD = V[V['NOTD']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND + BND + Partial)#Cost of SKUs not received by OTD

idx = ['Not on Brightpearl', 'Not Booked Not Delivered', 'Booked Not Delivered', 'Partial Delivery']
WorkingCapital = pd.DataFrame(data = [NoBP, NBND, BND, Partial], index = idx).T
WorkingCapital.applymap(lambda x: "R{:.8n}".format(x))
WorkingCapital.name = "Working Capital"

writer3 = ExcelWriter('04_Visibility\\ProductTrack QuickStats ' + str(today) + '.xlsx')
SimplesCount.to_excel(writer3, 'Sheet1', startrow = 3)
UnitsCount.to_excel(writer3, 'Sheet1', startrow = 10)
POCount.to_excel(writer3, 'Sheet1', startrow = 17)
WorkingCapital.to_excel(writer3, 'Sheet1', startrow = 24)
workbook = writer3.book
#format workbook
title = workbook.add_format({'bold':True, 'size':14})
header = workbook.add_format({'size':12, 'underline':True, 'font_color':'green'})
worksheet = writer3.sheets['Sheet1']
worksheet.write('A1','Spree Stock Tracking Statistics ' + str(today), title)
worksheet.write('A3','Simples Count (% of Simples Planned)', header)
worksheet.write('A11','Units Count', header)
worksheet.write('A18','PO Count', header)
worksheet.write('A25','Working Capital (ZAR loss due to status not achieved)', header)
worksheet.set_column('A:A', 8 )
worksheet.set_column('B:K', 18)
writer3.save()

#format QuickStats with openpyxl
wb = load_workbook('04_Visibility\\ProductTrack QuickStats ' + str(today) + '.xlsx')
ws = wb.worksheets[0]

cellsA = [ws['E4'],ws['E11'],ws['G4'],ws['D18'],ws['C25'],ws['D25']]
for cell in cellsA:
    cell.style.alignment.wrap_text = True 
    
cellsB = ws['B20':'F24']
for row in cellsB:
    for cell in row:
        cell.style.number_format.format_code = '0%'
        cell.style.alignment.horizontal = 'center'

cellsC = ws['B27':'F31']
for row in cellsC:
    for cell in row:
        cell.style.number_format.format_code = '"R "#,##0.00'
        
wb.save('04_Visibility\\ProductTrack QuickStats ' + str(today) + '.xlsx')

doc_name = 'ProductTrack QuickStats '
part = '04_Visibility\\ProductTrack QuickStats ' + str(today) + '.xlsx'
message = 'Where is my stock? Quick Stats to monitor production progress'
maillist = "MailList_QS.txt"

MyFunx.send_message(doc_name, message, part, maillist)

#==============================================================================
# Production Track Weekly Stats
#==============================================================================
POTrack = V.drop_duplicates(subset = ['POs'])
POTrack = POTrack.groupby(['Buyer','GLMonth','GLDay'])
POTrack = POTrack.apply(lambda x : pd.Series(dict(POTotal = x.POs.count(),PODraft = (x['Status'] == 'Draft PO').sum(),NOTReceived = (x['Date received'].isnull()).sum(),POReceived = (x['Date received'].notnull()).sum())))
POTrack = POTrack[['POTotal','PODraft','POReceived','NOTReceived']]
POTrack.name = "PO Track"

doc_name = 'GLTrack QuickStats '
part = '04_Visibility\\GoLiveTrack QuickStats ' + str(today) + '.xlsx'
message = 'Quick Stats to monitor Purchase Orders received to meet Go Live Dates'
maillist = "MailList_Prod.txt"

writer4 = ExcelWriter(part)
POTrack.to_excel(writer4, 'Sheet1', startrow = 2)
workbook = writer4.book
#format workbook
title = workbook.add_format({'bold':True, 'size':14})
header = workbook.add_format({'size':12, 'underline':True, 'font_color':'green'})
worksheet = writer4.sheets['Sheet1']
worksheet.write('A1','Spree Inbound POs to make Go Live Statistics ' + str(today), title)
#worksheet.write('A3','Simples Count (% of Simples Planned)', header)
worksheet.set_column('A:K', 12 )
writer4.save()

MyFunx.send_message(doc_name, message, part, maillist)