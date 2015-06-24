# -*- coding: utf-8 -*-
"""
Created on Tue Feb 10 11:12:34 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
from pandas import DataFrame
from datetime import date, datetime, timedelta
from pandas import ExcelWriter
import MyFunx, gdocs

shift = datetime.today().strftime('%Y-%m-%d %H:%M')
today = date.today() 
#==============================================================================
# Read Samples Plan master data
#==============================================================================

pw = "Spr33Pops101"
Lulu =  MyFunx.sql_import("vw_ProcurementPipeline","ActualGoLiveDate", pw)
Planned = Lulu[['PlannedGoLiveDayOfWeek','PlannedGoLiveMonth','PlannedGoLiveYear','BuyerPlanName','BuyerPlanStatus','EmployeeFirstName','PlannedUnitCostExclTax','PlannedTotalQuantity','PlannedTotalCostExclTax','SimpleSKU','SimpleName','ConfigName','ConfigSKU','ProcurementStatus','ProcurementProductCategoryL3','ActualGoLiveDate','Supplier','Designer','EANNumber','BarCode']]
Planned.rename(columns = {'PlannedGoLiveDayOfWeek':'GLDay','PlannedGoLiveMonth':'GLMonth','PlannedGoLiveYear':'GLYear','EmployeeFirstName':'Buyer','SimpleSKU':'ProductID'}, inplace = True)
Planned.drop_duplicates(subset = ['ProductID','GLMonth'], inplace = True, take_last = True)
Planned = Planned[Planned['PlannedTotalCostExclTax'] > 0]

Stock = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\02_StockCount\\Rolling Stock.xlsx')
QCed = Stock.parse('Sheet1', skiprows = 0, index = None, parse_cols = (0,1,4), parse_dates = True)
QCed.rename(columns={'Date': 'Date QCed'}, inplace=True)

SamplesPlan = pd.merge(Planned, QCed, on = 'ProductID', how = 'left', sort = False)
SamplesPlan.drop_duplicates(subset = ['ConfigSKU'],take_last = True, inplace = True)
SamplesPlan = SamplesPlan[(SamplesPlan.GLMonth >= today.month - 2) & (SamplesPlan.GLMonth <= today.month + 2) & (SamplesPlan.GLYear == today.year)]

#==============================================================================
# Read data from Google docs
#==============================================================================
#History Data
#data = pd.ExcelFile('05_Samples\\SampleTrack.xlsx')
#SH = data.parse('Master', header = 0, skiprows = 0, parse_cols = 'E,L:R', parse_dates = True)
columns = [u'SKU', u'01_WHSamples_OUT', u'02_SampleRoom_IN', u'03_SampleRoom_TO_studio', u'04_SampleRoom_FROM_studio', u'05_SampleRoom_OUT', u'06_WHSamples_IN']
SH = pd.read_csv('SampleTrack.txt',sep=';', header = 0, usecols = columns, parse_dates = [1,2,3,4,5,6], dayfirst=True)
SH.drop_duplicates(subset = ['SKU','01_WHSamples_OUT'], take_last = True, inplace = True)

c = gdocs.authenticate_gdocs()

sheets = pd.Series( ['01_WHSamples_OUT','02_SampleRoom_IN','03_SampleRoom_TO_studio','04_SampleRoom_FROM_studio','05_SampleRoom_OUT','06_WHSamples_IN'])

samples = pd.DataFrame(columns = ['SKU'])

x = range(1,7)
for i in x:
    sheet = sheets.ix[i-1]
    vals = SH.groupby(['SKU',sheet]).size()
    if vals.empty:
        SHi = pd.DataFrame(columns = ['SKU',sheet])
    else: 
        SHi = pd.DataFrame(data = vals)
        SHi.reset_index(inplace = True)
        SHi = SHi[['SKU',sheet]]
    
    sht = c.open(sheet)
    worksheet = sht.worksheet('TeamS')
    if worksheet.cell(2,1).value == "":
        print sheet, "is empty"
        count = pd.DataFrame(columns = ['SKU',sheet])
    else:
        info = worksheet.get_all_values()
        headers = info.pop(0)
        count = DataFrame(data = info, columns = ['SKU'])
        dup = [unicode(d.upper()) for d in count['SKU']]
        count['SKU'] = dup
        count[sheet] = shift
        count[sheet] = pd.to_datetime(count[sheet],coerce=True)
    
    PD = pd.concat([SHi, count], axis = 0, copy = False)
    samples = pd.merge(samples, PD, on = 'SKU', how = 'outer')
    
    l = len(count)        
    clean = worksheet.range('A2:A' + str(l + 5))
    for cl in clean:
        cl.value = ""
    worksheet.update_cells(clean)
    
samples['ConfigSKU'] = samples.SKU.apply(lambda x : x[:7] if len(x)==11 else x)

AllSamples = pd.merge(SamplesPlan, samples, how = 'outer', on = 'ConfigSKU', sort  = False)
AllSamples.drop_duplicates(subset = ['ConfigSKU','01_WHSamples_OUT'], take_last = True, inplace = True)
AllSamples = AllSamples.sort(columns = ['01_WHSamples_OUT'], ascending = True, na_position = 'last')
AllSamples = AllSamples[['GLDay','GLMonth','GLYear', 'ConfigSKU', 'SKU', 'ConfigName', 'Supplier', 'Designer','Category','Buyer','Date QCed','01_WHSamples_OUT','02_SampleRoom_IN','03_SampleRoom_TO_studio','04_SampleRoom_FROM_studio','05_SampleRoom_OUT','06_WHSamples_IN']]

SampleSummary = AllSamples.groupby(['GLYear','GLMonth']).agg({'Date QCed':'count','01_WHSamples_OUT':'count','02_SampleRoom_IN':'count','03_SampleRoom_TO_studio':'count','04_SampleRoom_FROM_studio':'count','05_SampleRoom_OUT':'count','06_WHSamples_IN':'count'})
SampleSummary = SampleSummary[['Date QCed','01_WHSamples_OUT','02_SampleRoom_IN','03_SampleRoom_TO_studio','04_SampleRoom_FROM_studio','05_SampleRoom_OUT','06_WHSamples_IN']]

doc_name = 'Samples Tracker '
part = '05_Samples\\SampleTrack ' + str(today) + '.xlsx'
message = 'Spree Samples Tracking ' + str(date.today())
maillist = "MailList_Samples.txt" 
     
writer = ExcelWriter(part)

SampleSummary.to_excel(writer,'EasyTrack',encoding = 'utf-8')
AllSamples.to_excel(writer,'Master', index = False, encoding = 'utf-8')
workbook = writer.book
wksht = writer.sheets['Master']

wksht.set_column('A:B', 10)
wksht.set_column('C:C', 14)
wksht.set_column('D:G', 18)
wksht.set_column('H:K', 10)
wksht.set_column('L:L', 18)
wksht.set_column('M:R', 28)
wksht = writer.sheets['EasyTrack']
wksht.set_column('A:C', 8)
wksht.set_column('D:D', 16)
wksht.set_column('E:J', 29)
writer.save()

MyFunx.send_message(doc_name, message, part, maillist)  
    
#Create SampleTrack Reference doc
    
#part2 = '05_Samples\\SampleTrack.xlsx'
     
#writer2 = ExcelWriter(part2)
#AllSamples.to_excel(writer2,'Master', index = False)
#workbook2 = writer2.book
#wksht2 = writer2.sheets['Master']
#wksht2.set_column('A:B', 10)
#wksht2.set_column('C:C', 14)
#wksht2.set_column('D:G', 18)
#wksht2.set_column('H:K', 10)
#wksht2.set_column('L:L', 18)
#wksht2.set_column('M:R', 28)
#writer2.save()

AllSamples.to_csv('SampleTrack.txt', sep=';',index=False, encoding = 'utf-8')