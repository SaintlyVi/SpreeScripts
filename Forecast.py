# -*- coding: utf-8 -*-
"""
Created on Fri Mar 20 11:14:26 2015

@author: wiebke.toussaint
"""

import numpy as np
from pandas import ExcelWriter, pivot_table
from datetime import date, timedelta
import MyFunx, AllData

def iso_year_start(iso_year):
    "The gregorian calendar date of the first day of the given ISO year"
    fourth_jan = date(iso_year, 1, 4)
    delta = timedelta(fourth_jan.isoweekday()-1)
    return fourth_jan - delta 

def iso_to_gregorian(iso_year, iso_week, iso_day):
    "Gregorian calendar date for the given ISO year, week and day"
    year_start = iso_year_start(iso_year)
    return year_start + timedelta(days=iso_day-1, weeks=iso_week-1)

#from pandas import ExcelWriter

today = date.today()
past = today.month - 2
future = today.month + 2

lastmonth = past
nextmonth = future

Visibility = AllData.InboundData(lastmonth, nextmonth)
V1 = Visibility[Visibility['Ref'].str.contains("sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud")==False] 
V2 = Visibility[Visibility['Ref'].isnull()==True]
V = V1.append(V2, ignore_index=True)
V = V.sort(['Date received','Date booked','POs'], inplace = False, na_position = 'first')
V = V[['GLDate','Buyer', 'UnitCost','TotalUnits','TotalCost','SKU','SimpleName','ProcurementStatus','Category','Supplier','DeliveryDue','POs','BP Qty','Ref','Status','Date booked','Partial delivery','Date received','LastQCed','Qty Counted','Qty Damaged','OTDLastReceived','Qty Received','Qty PutAway','ActualGoLiveDate']]
    
V.loc[V['DeliveryDue'].isnull(),'DeliveryDue'] = V['GLDate'] - timedelta(14)
V['DeliveryDue'] = V['DeliveryDue'].apply(lambda x: list(x.isocalendar())[:-1] + [3])
V['DeliveryWeek'] = V['DeliveryDue'].apply(lambda x: iso_to_gregorian(x[0],x[1],x[2]))
V['DeliveryMonth'] = V['DeliveryWeek'].apply(lambda x: x.month)
V['DeliveryYear'] = V['DeliveryWeek'].apply(lambda x: x.year)
V = V[V.ProcurementStatus != 'Deleted']

if past == 11 | 12:
    V = V[((V.DeliveryMonth >= past) & (V.DeliveryYear == today.year)) | ((V.DeliveryMonth <= future) & (V.DeliveryYear == today.year))]
else:    
    V = V[((V.DeliveryMonth >= past) & (V.DeliveryYear == today.year)) & ((V.DeliveryMonth <= future) & (V.DeliveryYear == today.year))]

table = pivot_table(V, values = 'TotalUnits', index = 'Category', columns = 'DeliveryWeek', aggfunc = np.sum, dropna = True, fill_value = "", margins = True)

received = V.groupby('DeliveryWeek').aggregate({'TotalUnits':np.sum, 'Qty Received':np.sum})
received.columns = 'Qty Received', 'Qty Planned'

writer = ExcelWriter("07_Forecast\\Forecast " + str(today.year) + "-" + str(today.month) + ".xlsx")
received.to_excel(writer, 'Summary', index = True)
table.to_excel(writer, 'Config Plan', index = True)
workbook = writer.book
wksht = writer.sheets['Config Plan']
wksht.set_column('A:A', 25)
wksht.set_column('B:V', 10)
ws = writer.sheets['Summary']
ws.set_column('A:C', 15)
writer.save()
