# -*- coding: utf-8 -*-
"""
Created on Wed Apr 01 10:41:56 2015

@author: Wiebke.Toussaint
"""

import numpy as np
from pandas import ExcelWriter
from datetime import date
import AllData# , MyFunx

today = date.today()

lastmonth = today.month - 3
nextmonth = today.month + 1

Visibility = AllData.InboundData(lastmonth, nextmonth, today)
V1 = Visibility[Visibility['Ref'].str.contains("sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud")==False] 
V2 = Visibility[Visibility['Ref'].isnull()==True]
V = V1.append(V2, ignore_index=True)
V = V.sort(['Date received','Date booked','POs'], inplace = False, na_position = 'first')
V = V[['GLYear','GLMonth','GLDay','Buyer', 'UnitCost','TotalUnits','TotalCost','SKU','SimpleName','ProcurementStatus','Category','Supplier','DeliveryDue','POs','BP Qty','Ref','Status','Date booked','Partial delivery','Date received','LastQCed','Qty Counted','Qty Damaged','OTDLastReceived','Qty Received','Qty PutAway','ActualGoLiveDate']]

#==============================================================================
# Generate Supplier Compliance
#==============================================================================

SC = V[['GLMonth','Supplier','POs','Status','SKU','TotalUnits','TotalCost','DeliveryDue','Date booked','Date received','Partial delivery','Qty Counted','Qty Damaged','Buyer']]

SC = SC[(SC['Status'].notnull() == True) & (SC['Status'] != 'Draft PO') & (SC['TotalUnits'] != 0)]

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

writerx = ExcelWriter('06_Compliance\\SupplierCompliance 2015-' + str(today.month - 1) + '.xlsx')
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

writerx = ExcelWriter('06_Compliance\\BuyerCompliance 2015-' + str(today.month - 1) + '.xlsx')
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


