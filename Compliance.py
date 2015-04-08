# -*- coding: utf-8 -*-
"""
Created on Wed Apr 01 10:41:56 2015

@author: Wiebke.Toussaint
"""

import numpy as np
import pandas as pd
from pandas import DataFrame, ExcelWriter
from datetime import date
import gspread
import MyFunx

today = date.today()
lastmonth = today.month - 2
nextmonth = today.month + 1

#==============================================================================
# Import from all required data sources
#==============================================================================
#Import Brightpearl Detail Report data
columns = ["Order ID", "Ref", "SKU", "Status", "Quantity"]
BPdetail = pd.read_csv('BPdetail.csv', header = 0, usecols = columns)
BPdetail['Order ID'] = BPdetail['Order ID'].map(lambda x: str(x))
BPdetail.rename(columns={'Order ID': 'POs', 'Quantity':'BP Qty'}, inplace=True)
BPdet = BPdetail[BPdetail['Status'].str.contains('Cancel PO')==False] 

CancelledPOs = BPdetail[BPdetail['Status']=='Cancel PO']
CancelledPOs = CancelledPOs.groupby('POs').agg({'SKU':'count'})

#Import Brightpearl PO Report data
columns = ["Order ID", "Delivery due"]
BPreport = pd.read_csv('BPreport.csv', header = 0, usecols = columns, parse_dates = [1])
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
Damages = pd.ExcelFile('03_Damages_OS\\Rolling Damages.xlsx')
Damages = Damages.parse('Sheet1', skiprows = 0, index = None)
SKU = Damages['ProductID'].value_counts()
Damagd = DataFrame(data = SKU)
Damagd.reset_index(level=0, inplace=True)
Damagd.columns = ['SKU', 'Qty Damaged']

#Import Lulu Assortment Plans
## table = "vw_ProcurementPipeline"
## dateparse = "ActualGoLiveDate"

pw = raw_input("Enter SQL Server database password: ")
Lulu =  MyFunx.sql_import("vw_ProcurementPipeline","ActualGoLiveDate",pw)
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
Planned = Planned[Planned['ProcurementStatus'] != 'Deleted']

#Import Dynaman IBOI1003 Inbound Order Received Messages
## table = "vw_WarehouseInboundItemsReceived"
## dateparse = "Timestamp"

IBOI1003 =  MyFunx.sql_import("vw_WarehouseInboundItemsReceived","Timestamp",pw)
IBOI1003 = IBOI1003[['MessageReference','ItemCode','QuantityReceived','Timestamp']]
TaknIn = IBOI1003.groupby(['MessageReference','ItemCode']).agg({'QuantityReceived':np.sum, 'Timestamp':np.max})
TaknIn.reset_index(inplace=True)
TaknIn.columns = ['POs','SKU','OTDLastReceived','Qty Received']
TaknIn['POs'] = TaknIn['POs'].apply(lambda x: x if len(x) < 7 else 0)
TaknIn = TaknIn[TaknIn['POs'] != 0]

#Import Dynaman ITMI1002 
## table = "vw_WarehouseStockAvailability"
## dateparse = "Timestamp"

ITMI1002 =  MyFunx.sql_import("vw_WarehouseStockAvailability","Timestamp",pw)
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
# Generate Supplier Compliance
#==============================================================================

SC = V[['GLMonth','Supplier','POs','Status','SKU','TotalUnits','TotalCost','DeliveryDue','Date booked','Date received','Partial delivery','Qty Counted','Qty Damaged','Buyer']]

SC = SC[(SC['Status'].notnull() == True) & (SC['Status'] != 'Draft PO')]

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

SupCompMonth = SupComp[SupComp.GLMonth == today.month - 1]

writerx = ExcelWriter('SupplierCompliance 2015-' + str(today.month - 1) + '.xlsx')
SupCompMonth.to_excel(writerx, 'Last Month', index = False )
SupComp.to_excel(writerx, 'SupComp', index = False)   
workbook = writerx.book
worksheet = writerx.sheets['SupComp']
worksheet.set_column('A:A', 8 )
worksheet.set_column('B:B', 25)
worksheet.set_column('C:E', 12)
worksheet.set_column('F:I', 22)
worksheet.set_column('J:K', 10)
worksht = writerx.sheets['Last Month']
worksht.set_column('A:A', 8 )
worksht.set_column('B:B', 25)
worksht.set_column('C:E', 12)
worksht.set_column('F:I', 22)
worksht.set_column('J:K', 10)
writerx.save()
    
#==============================================================================
# Generate Buyer Compliance data
#==============================================================================
#calculating by month
bc0 = SC.groupby(['GLMonth','Buyer','POs']).agg({'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()

bc1 = bc0.groupby(['GLMonth','Buyer']).agg({'POs':np.size,'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()

bc1['InFull'] = abs(bc1.apply(lambda x : x['Qty Counted'] / x['TotalUnits'], axis = 1))

bc0['int1'] = bc0['MissedBooking_days'].replace(np.nan, -1)
bc0.loc[bc0['int1']>0,'int1'] = 0
bc1['NotBooked'] = abs(bc0.groupby(['GLMonth','Buyer'])['int1'].sum().reset_index()['int1'])
bc1.NotBooked = bc1.NotBooked / bc1.POs

bc0['int2'] = bc0['MissedPlan_days'].replace(np.nan, -1)
bc0.loc[bc0['int2']>0,'int2'] = 0
bc1['NotDelivered'] = abs(bc0.groupby(['GLMonth','Buyer'])['int2'].sum().reset_index()['int2'])
bc1.NotDelivered = bc1.NotDelivered / bc1.POs

BuyerComp = bc1[['GLMonth','Buyer','TotalCost','POs','TotalUnits','MissedBooking_days','NotBooked','MissedPlan_days','NotDelivered','InFull','Qty Damaged']]
BuyerComp = BuyerComp.rename(columns={'MissedBooking_days':'Missed booking (avg days)','MissedPlan_days':'Missed DeliveryDue (avg days)'}, inplace = False)
BuyerComp = BuyerComp.sort(columns = ['GLMonth','NotBooked','NotDelivered','InFull','TotalCost'], ascending = [1,0,0,1,0], inplace = False, na_position = 'first')

#calculating summary for period
sum0 = SC.groupby(['Buyer','POs']).agg({'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()

sum1 = sum0.groupby(['Buyer']).agg({'POs':np.size,'TotalUnits':np.sum,'TotalCost':np.sum,'Qty Counted':np.sum,'Qty Damaged':np.sum,'MissedBooking_days':np.mean,'MissedPlan_days':np.mean}).reset_index()

sum1['InFull'] = abs(sum1.apply(lambda x : x['Qty Counted'] / x['TotalUnits'], axis = 1))

sum0['int1'] = sum0['MissedBooking_days'].replace(np.nan, -1)
sum0.loc[sum0['int1']>0,'int1'] = 0
sum1['NotBooked'] = abs(sum0.groupby('Buyer')['int1'].sum().reset_index()['int1'])
sum1.NotBooked = sum1.NotBooked / sum1.POs

sum0['int2'] = sum0['MissedPlan_days'].replace(np.nan, -1)
sum0.loc[sum0['int2']>0,'int2'] = 0
sum1['NotDelivered'] = abs(sum0.groupby('Buyer')['int2'].sum().reset_index()['int2'])
sum1.NotDelivered = sum1.NotDelivered / sum1.POs

BuCompSum = sum1[['Buyer','TotalCost','POs','TotalUnits','MissedBooking_days','NotBooked','MissedPlan_days','NotDelivered','InFull','Qty Damaged']]
BuCompSum = BuCompSum.rename(columns={'MissedBooking_days':'Missed booking (avg days)','MissedPlan_days':'Missed DeliveryDue (avg days)'}, inplace = False)
BuCompSum = BuCompSum.sort(columns = ['NotBooked','NotDelivered','InFull','TotalCost'], ascending = [0,0,1,0], inplace = False, na_position = 'first')

BuCompMonth = BuyerComp[BuyerComp.GLMonth == today.month - 1]

writerx = ExcelWriter('BuyerCompliance 2015-' + str(today.month - 1) + '.xlsx')
BuCompMonth.to_excel(writerx, 'Last Month', index = False)   
BuyerComp.to_excel(writerx, 'BuyerComp', index = False) 
BuCompSum.to_excel(writerx, '2015 Summary', index = False)  
workbook = writerx.book
worksheet = writerx.sheets['BuyerComp']
worksheet.set_column('A:A', 8 )
worksheet.set_column('B:B', 25)
worksheet.set_column('C:E', 12)
worksheet.set_column('F:I', 22)
worksheet.set_column('J:K', 10)
wsheet = writerx.sheets['Last Month']
wsheet.set_column('A:A', 8 )
wsheet.set_column('B:B', 25)
wsheet.set_column('C:E', 12)
wsheet.set_column('F:I', 22)
wsheet.set_column('J:K', 10)
ws = writerx.sheets['2015 Summary']
ws.set_column('A:B', 14 )
ws.set_column('C:D', 12)
ws.set_column('E:H', 22)
ws.set_column('J:K', 10)
writerx.save()


