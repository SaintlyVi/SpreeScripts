#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 02 12:53:35 2014

@author: Wiebke.Toussaint
"""
#==============================================================================
# #VISIBILITY
# #Product Tracking Report looking 5 weeks back, 1 week forward
# #TO DO before running the script:
# #1. Download -5week + 1 week Brightpearl Detail Report (filter DELIVERY DATE)
# #2. Download -2month + 1 month Brightpearl PO Report
# #3. Refresh Lulu PowerQuery
# #4. Refresh IBOI1003 PowerQuery
# #5. Run Damages Script
# #6. Run Stock Count Script
#==============================================================================
import sys
sys.path.append('Z:\\SUPPLY CHAIN\\Python Scripts\\00_SharedFunctions')

import numpy as np
import pandas as pd
from pandas import DataFrame, ExcelWriter
from datetime import date, timedelta
import gspread
from openpyxl.reader.excel import load_workbook
from dateutil.parser import parse
import Email, SQL
#from pandas import ExcelWriter

today = date.today()
lastmonth = today.month - 2
nextmonth = today.month + 1

#==============================================================================
# Import from all required data sources
#==============================================================================
#Import Brightpearl Detail Report data
columns = ["Order ID", "Ref", "SKU", "Status", "Quantity"]
BPdetail = pd.read_csv('Z:\\SUPPLY CHAIN\\Python Scripts\\00_UPDATE\\BPdetail.csv', header = 0, usecols = columns)
BPdetail['Order ID'] = BPdetail['Order ID'].map(lambda x: str(x))
BPdetail.rename(columns={'Order ID': 'POs', 'Quantity':'BP Qty'}, inplace=True)
BPdet = BPdetail[BPdetail['Status'].str.contains('Cancel PO')==False] 

CancelledPOs = BPdetail[BPdetail['Status']=='Cancel PO']
CancelledPOs = CancelledPOs.groupby('POs').agg({'SKU':'count'})

#Import Brightpearl PO Report data
columns = ["Order ID", "Delivery due"]
BPreport = pd.read_csv('Z:\\SUPPLY CHAIN\\Python Scripts\\00_UPDATE\\BPreport.csv', header = 0, usecols = columns, parse_dates = [1])
BPreport = BPreport.dropna(axis = 0,how = 'all') #removes empty rows
BPreport['Order ID'] = BPreport['Order ID'].map(lambda x: x.strip('PO#')) #removes text in front of PO number
BPreport.rename(columns={'Order ID': 'POs', 'Delivery due':'DeliveryDue'}, inplace=True)

BP = pd.merge(BPdet, BPreport, on = 'POs', how = 'left', sort = False)

#Import Epping Receiving Report data
c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
c.login()

sht = c.open('Epping Receiving Report')
worksheet = sht.worksheet('Booked')

info = worksheet.get_all_values()
headers = info.pop(0)
B_R = DataFrame(data = info, columns = headers)

Bookd = B_R[['POs','Date booked']]
Bookd = Bookd.replace('',np.nan)
Bookd = Bookd.dropna(subset = ['Date booked'], thresh = 1)
Bookd = Bookd.drop_duplicates(subset = ['POs'], take_last = False)
Bookd['Date booked'] = pd.to_datetime(Bookd['Date booked'], infer_datetime_format = True)

Receivd = B_R[['POs', 'Partial delivery', 'Date received']]
Receivd = Receivd.replace('',np.nan)
Receivd = Receivd.dropna(subset = ['Date received'], thresh = 1)
Receivd = Receivd.drop_duplicates(subset = ['POs'], take_last = True)
Receivd = Receivd[Receivd.POs != np.nan]
Receivd['Date received'] = pd.to_datetime(Receivd['Date received'], infer_datetime_format = True)

#Import Rolling Stock data
Stock = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\02_StockCount\\Rolling Stock.xlsx')
QCed = Stock.parse('Sheet1', skiprows = 0, index = None, parse_cols = (1,3,4,5))
QCed.rename(columns={'Date': 'LastQCed', 'PO':'POs','ProductID':'SKU'}, inplace=True)
poqc = [str(p) for p in QCed['POs']]
QCed['POs'] = poqc
QCed = QCed.groupby(['POs','SKU']).agg({'Qty Counted':np.sum, 'LastQCed':np.max})
QCed.reset_index(inplace=True)

#Import Rolling Damages
Damages = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\03_Damages\\Rolling Damages.xlsx')
Damages = Damages.parse('Sheet1', skiprows = 0, index = None)
SKU = Damages['SKU'].value_counts()
Damagd = DataFrame(data = SKU)
Damagd.reset_index(level=0, inplace=True)
Damagd.columns = ['SKU', 'Qty Damaged']

#Import Lulu Assortment Plans
#table = "vw_ProcurementPipeline"
#dateparse = "ActualGoLiveDate"

Lulu =  SQL.sql_import("vw_ProcurementPipeline","ActualGoLiveDate")
Planned = Lulu[['PlannedGoLiveDayOfWeek','PlannedGoLiveMonth','PlannedGoLiveYear','BuyerPlanName','BuyerPlanStatus','EmployeeFirstName','PlannedUnitCostExclTax','PlannedTotalQuantity','PlannedTotalCostExclTax','SimpleSKU','SimpleName','ConfigName','ConfigSKU','ProcurementStatus','ProcurementProductCategoryL3','ActualGoLiveDate','Supplier','Designer','EANNumber','BarCode']]
#Merge EAN, BarCode information with SKU
SKU = Planned['EANNumber'].combine_first(Planned['SimpleSKU'])
Planned['SKU'] = Planned['BarCode'].combine_first(SKU)
Planned.drop_duplicates(subset = ['SKU','PlannedGoLiveMonth'], inplace = True, take_last = True)
if lastmonth == 11 | 12:
    Planned = Planned[((Planned.PlannedGoLiveMonth >= lastmonth) & (Planned.PlannedGoLiveYear == today.year)) | ((Planned.PlannedGoLiveMonth <= nextmonth) & (Planned.PlannedGoLiveYear == today.year))]
else:    
    Planned = Planned[((Planned.PlannedGoLiveMonth >= lastmonth) & (Planned.PlannedGoLiveYear == today.year)) & ((Planned.PlannedGoLiveMonth <= nextmonth) & (Planned.PlannedGoLiveYear == today.year))]
Planned.rename(columns={'PlannedGoLiveDayOfWeek':'GLDay','PlannedGoLiveMonth':'GLMonth','PlannedGoLiveYear':'GLYear', 'EmployeeFirstName':'Buyer','ProcurementProductCategoryL3':'Category', 'PlannedUnitCostExclTax':'UnitCost','PlannedTotalQuantity':'TotalUnits','PlannedTotalCostExclTax':'TotalCost'}, inplace=True)
Planned = Planned[Planned['TotalCost'] > 0]

#Import Dynaman IBOI1003 Inbound Order Received Messages
#table = "vw_WarehouseInboundItemsReceived"
#dateparse = "Timestamp"

IBOI1003 =  SQL.sql_import("vw_WarehouseInboundItemsReceived","Timestamp")
IBOI1003 = IBOI1003[['MessageReference','ItemCode','QuantityReceived','Timestamp']]
TaknIn = IBOI1003.groupby(['MessageReference','ItemCode']).agg({'QuantityReceived':np.sum, 'Timestamp':np.max})
TaknIn.reset_index(inplace=True)
TaknIn.columns = ['POs','SKU','OTDLastReceived','Qty Received']
TaknIn['POs'] = TaknIn['POs'].apply(lambda x: x if len(x) < 7 else 0)
TaknIn = TaknIn[TaknIn['POs'] != 0]

#Import Dynaman ITMI1002 
#table = "vw_WarehouseStockAvailability"
#dateparse = "Timestamp"

ITMI1002 =  SQL.sql_import("vw_WarehouseStockAvailability","Timestamp")
ITMI1002 = ITMI1002[['ITEM_CODE','QTY']]
PutAway = pd.pivot_table(ITMI1002, values = ['QTY'], index = ['ITEM_CODE'], aggfunc=np.sum)
PutAway.reset_index(inplace=True)
PutAway.columns = ['SKU','Qty PutAway']

Merge = pd.merge(Planned, BP, on = 'SKU', how = 'left', sort = False)
Merge1 = pd.merge(Merge, Bookd, on = 'POs', how = 'left', sort = False)
Merge2 = pd.merge(Merge1, Receivd, on = 'POs', how = 'left', sort = False)
Merge3 = pd.merge(Merge2, QCed, on = ['SKU','POs'], how = 'left', sort = False)
Merge4 = pd.merge(Merge3, Damagd, on = 'SKU', how = 'left', sort = False)
Merge5 = pd.merge(Merge4, TaknIn, on = ['SKU','POs'], how = 'left', sort = False)
Visibility = pd.merge(Merge5, PutAway, on = 'SKU', how = 'left')
Visibility.drop_duplicates(inplace = True)
Va = Visibility[Visibility.duplicated(subset = ['SKU'], take_last = True)==False]
Vb = Visibility[Visibility.duplicated(subset = ['SKU'], take_last = True)==True]
Vc = Vb[Vb['Status']!='Draft PO']
Visibility = Va.append(Vc, ignore_index = True)
Visibility.replace("", np.nan, inplace = True)

V1 = Visibility[Visibility['Ref'].str.contains("sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud")==False] 
V2 = Visibility[Visibility['Ref'].isnull()==True]
V = V1.append(V2, ignore_index=True)
V = V.sort(['Date received','Date booked','POs'], inplace = False, na_position = 'first')
V = V[['GLYear','GLMonth','GLDay','Buyer', 'UnitCost','TotalUnits','TotalCost','SKU','SimpleName','ProcurementStatus','Category','Supplier','DeliveryDue','POs','BP Qty','Ref','Status','Date booked','Partial delivery','Date received','LastQCed','Qty Counted','Qty Damaged','OTDLastReceived','Qty Received','Qty PutAway','ActualGoLiveDate']]

OS = Visibility[Visibility['Ref'].str.contains("OS|Os|OVERSUPPLY")==True]
OS = OS[['GLYear','GLMonth','GLDay','Buyer', 'UnitCost','TotalUnits','TotalCost','SKU','SimpleName','Category','Supplier','POs','Ref','Status','Qty Damaged','OTDLastReceived','Qty Received']]

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
    worksheet.set_column('V:W', 12 )
    worksheet.set_column('X:X', 18 )
    worksheet.set_column('Y:Z', 12 )
    worksheet.set_column('AA:AA', 18 )

writer1 = ExcelWriter('Visibility ' + str(today) + '.xlsx')
V.to_excel(writer1, 'MASTER', index = False )   
workbook = writer1.book
worksheet = writer1.sheets['MASTER']
format()
Buyers = pd.Series(V['Buyer'].unique())
for b in Buyers:
    DataName = V[V['Buyer']==b]
    DataName.to_excel(writer1, b , index = False )   
    workbook = writer1.book
    worksheet = writer1.sheets[b]
    format()
OS.to_excel(writer1, 'OS', index = False )
workbook = writer1.book
worksheet = writer1.sheets['OS']
writer1.save()

doc_name = 'Visibility Report '
part = 'Visibility ' + str(today) + '.xlsx'
message = 'Visibility Report' + str(today)
maillist = "MerchMailList.txt"

Email.send_message(doc_name, message, part, maillist)

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

writer2 = ExcelWriter('WHTrack ' + str(today) + '.xlsx')
Track = [WHTrack,Backlog]
for t in Track:
    t.to_excel(writer2, t.name , index = False )   
    workbook = writer2.book
    worksheet = writer2.sheets[t.name]
    format1()  
writer2.save()
   
doc_name = 'WH Stock Track Report '
part = 'WHTrack ' + str(today) + '.xlsx'
message = 'Spree Stock Tracking on ' + str(today)
maillist = "WHMailList.txt"

Email.send_message(doc_name, message, part, maillist)

#==============================================================================
# Generate ProductTrack QuickStats
#==============================================================================

#ProductTrack
SimplesCount = V.groupby('GLMonth').agg({'TotalUnits':'count','POs': 'count',\
'Date booked':'count', 'Date received':'count','Qty Counted':'count', 'Qty Received':'count','Qty PutAway':'count'})
SimplesCount.sort_index(ascending = True, inplace = True)
SimplesCount = SimplesCount[['TotalUnits','POs', 'Date booked', 'Date received','Qty Counted','Qty Received','Qty PutAway']]
SimplesCount.rename(columns={'TotalUnits':'Simples Planned','POs':'Simples on BP','Date booked':'Simples booked',\
'Date received':'Simples received UNCHECKED','Qty Counted':'Simples QCed','Qty Received':'Simples taken in by OTD','Qty PutAway':'Simples in OTD WH'}, inplace = True)
SimplesCount.name = "Simples Count"

UnitsCount = V.groupby('GLMonth').agg({'TotalUnits':'sum','Qty Counted':'sum', 'Qty Damaged':'sum','Qty Received':'sum','Qty PutAway':'sum'})
UnitsCount.sort_index(ascending = True, inplace = True)
UnitsCount = UnitsCount[['TotalUnits','Qty Counted','Qty Damaged','Qty Received','Qty PutAway']]
UnitsCount.rename(columns={'TotalUnits':'Units Planned','Qty Counted':'Units QCed','Qty Received':'Units taken in by OTD','Qty PutAway':'Units in OTD WH'}, inplace = True)
UnitsCount.name = "Units Count"

POCount = V.drop_duplicates(subset = ['POs'])
POCount = POCount.groupby('GLMonth').agg({'POs':'count', 'Date booked':'count','Date received':'count','LastQCed':'count', 'OTDLastReceived':'count'})
POCount.sort_index(ascending = True, inplace = True)
POCount = POCount[['POs', 'Date booked', 'Date received','LastQCed','OTDLastReceived']]
POCount.rename(columns={'POs':'POs on BP','Date booked':'POs booked','Date received':'POs received UNCHECKED','LastQCed':'POs QCed','OTDLastReceived':'POs in WH'}, inplace = True)
POCount = (POCount.T/list(POCount['POs on BP'])).T
POCount.name = "PO Count"

#SKUs not processed
V['TotalCost'] = V['UnitCost']*V['TotalUnits']
V['NoBP'] = V['Status'].isnull() #SKUs not on Brightpearl
V['NBND'] = V['Date received'].isnull() & V['Date booked'].isnull() #SKUs not booked not delivered
V['ND'] = V['Date received'].isnull() #SKUs not delivered
V['NQC'] = V['LastQCed'].isnull() #SKUs not QCed
V['NOTD'] = V['OTDLastReceived'].isnull() #SKUs not received by OTD

#Cost of SKUs not processed / WorkingCapital
NoBP = V[V['NoBP']==True].groupby(V['GLMonth']).sum()['TotalCost'] #Cost of SKUs not on BP / month
NBND = V[V['NBND']==True].groupby(V['GLMonth']).sum()['TotalCost'] - NoBP #Cost of SKUs not booked not delivered
BND = V[V['ND']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND) #Cost of SKUs not delivered / month
NQC = V[V['NQC']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND + BND)  #Cost of SKUs not QCed
NOTD = V[V['NOTD']==True].groupby(V['GLMonth']).sum()['TotalCost'] - (NoBP + NBND + BND + NQC)#Cost of SKUs not received by OTD

idx = ['Not on Brightpearl', 'Not Booked Not Delivered', 'Booked Not Delivered', 'Not QCed', 'OTD Not Received']
WorkingCapital = pd.DataFrame(data = [NoBP, NBND, BND, NQC, NOTD], index = idx).T
WorkingCapital.applymap(lambda x: "R{:.8n}".format(x))
WorkingCapital.name = "Working Capital"

writer3 = ExcelWriter('ProductTrack QuickStats ' + str(today) + '.xlsx')
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
wb = load_workbook('ProductTrack QuickStats ' + str(today) + '.xlsx')
ws = wb.worksheets[0]

cellsA = [ws['E4'],ws['E11'],ws['G4'],ws['D18'],ws['C25'],ws['D25']]
for cell in cellsA:
    cell.style.alignment.wrap_text = True 
    
cellsB = ws['B20':'F23']
for row in cellsB:
    for cell in row:
        cell.style.number_format.format_code = '0%'
        cell.style.alignment.horizontal = 'center'

cellsC = ws['B27':'F29']
for row in cellsC:
    for cell in row:
        cell.style.number_format.format_code = '"R "#,##0.00'
        
wb.save('ProductTrack QuickStats ' + str(today) + '.xlsx')

doc_name = 'ProductTrack QuickStats '
part = 'ProductTrack QuickStats ' + str(today) + '.xlsx'
message = 'Where is my stock? Quick Stats to monitor production progress'
maillist = "QSMailList.txt"

#if today.weekday() == 4:
Email.send_message(doc_name, message, part, maillist)

#==============================================================================
# Generate Supplier Compliance
#==============================================================================

if today.weekday() == 4:
    SC = V[['GLMonth','Supplier','POs','Status','SKU','TotalUnits','TotalCost','DeliveryDue','Date booked','Date received','Partial delivery','Qty Counted','Qty Damaged','Buyer']]
    
    SC = SC[(SC['Status'].notnull() == True) & (SC['Status'] != 'Draft PO')]
    
    #SC['DeliveryBookedYN'] = (SC['Date booked'].notnull() == True) & (SC['Qty Counted'].notnull() == True)
    
    #SC['Date booked'] = SC['Date booked'].apply(lambda x: parse(x) if type(x) is str else np.nan)
    #SC['Date received'] = SC['Date received'].apply(lambda x: parse(x) if type(x) is str else np.nan)
    #SC['DeliveryDue'] = SC['DeliveryDue'].apply(lambda x: x.to_datetime())
    
    totalseconds = 3600*24
    SC['MissedBooking_days'] = abs(SC['Date received'] - SC['Date booked'])/np.timedelta64(1,'D')
    
    SC['MissedPlan_days'] = abs(SC['DeliveryDue'] - SC['Date received'])/np.timedelta64(1,'D')
    
    sc0 = SC.groupby(['GLMonth','Supplier','POs']).agg({'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()
    
    sc1 = sc0.groupby(['GLMonth','Supplier']).agg({'POs':np.size,'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()
    
    sc1['InFull'] = abs(sc1.apply(lambda x : x['Qty Counted'] / x['TotalUnits'], axis = 1))
    
    sc0['int1'] = sc0['MissedBooking_days'].replace(np.nan, -1)
    sc0.loc[sc0['int1']>0,'int1'] = 0
    sc1['NotBooked'] = abs(sc0.groupby(['GLMonth','Supplier'])['int1'].sum().reset_index()['int1'])
    sc1.NotBooked = sc1.NotBooked / sc1.POs
    
    sc0['int2'] = sc0['MissedPlan_days'].replace(np.nan, -1)
    sc0.loc[sc0['int2']>0,'int2'] = 0
    sc1['NotDelivered'] = abs(sc0.groupby(['GLMonth','Supplier'])['int2'].sum().reset_index()['int2'])
    sc1.NotDelivered = sc1.NotDelivered / sc1.POs
    
    SupComp = sc1[['GLMonth','Supplier','TotalCost','POs','TotalUnits','MissedBooking_days','NotBooked','MissedPlan_days','NotDelivered','InFull','Qty Damaged']]
    SupComp = SupComp.rename(columns={'MissedBooking_days':'Missed booking (avg days)','MissedPlan_days':'Missed DeliveryDue (avg days)'}, inplace = False)
    SupComp = SupComp.sort(columns = ['GLMonth','NotBooked','NotDelivered','InFull','TotalCost'], ascending = [1,0,0,1,0], inplace = False, na_position = 'first')
    
    writerx = ExcelWriter('SupplierCompliance ' + str(today) + '.xlsx')
    SupComp.to_excel(writerx, 'SupComp', index = False)   
    workbook = writerx.book
    worksheet = writerx.sheets['SupComp']
    worksheet.set_column('A:A', 8 )
    worksheet.set_column('B:B', 25)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:I', 22)
    worksheet.set_column('J:K', 10)
    writerx.save()
    
