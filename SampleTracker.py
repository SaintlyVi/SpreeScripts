# -*- coding: utf-8 -*-
"""
Created on Tue Feb 10 11:12:34 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
from pandas import DataFrame
from datetime import date, datetime#, timedelta
import gspread
from pandas import ExcelWriter
import pyodbc
import pandas.io.sql as psql
import MyFunx

shift = datetime.today().strftime('%Y-%m-%d %H:%M')
today = date.today() 

#==============================================================================
# Read Samples Plan master data
#==============================================================================

connection_string1 = "DRIVER={SQL Server};SERVER=02CPT-TLSQL01;DATABASE=Spree SSBI;UID=SSBI_PaymentOps;PWD=Spr33Pops101;TABLE=vw_ProcurementPipeline"
cnxn1 = pyodbc.connect(connection_string1)
cursor1 = cnxn1.cursor()
sql1 = "select * from vw_ProcurementPipeline"  
df1 = psql.read_sql(sql1, cnxn1, parse_dates = ['ActualGoLiveDate'])
Planned = df1[['PlannedGoLiveDayOfWeek','PlannedGoLiveMonth','PlannedGoLiveYear','BuyerPlanName','BuyerPlanStatus','EmployeeFirstName','PlannedUnitCostExclTax','PlannedTotalQuantity','PlannedTotalCostExclTax','SimpleSKU','SimpleName','ConfigName','ConfigSKU','ProcurementStatus','ProcurementProductCategoryL3','ActualGoLiveDate','Supplier','Designer','EANNumber','BarCode']]

#Lulu1 = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\00_UPDATE\\Lulu1.xlsx')
#Planned = Lulu1.parse('Sheet4', skiprows = 0, index = None, parse_cols = (13,14,19,26,36,39,50,60))
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
data = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\06_Samples\\SampleTrack.xlsx')
SH = data.parse('Master', header = 0, skiprows = 0, parse_cols = 'C,M:R', parse_dates = True)
SH.drop_duplicates(inplace = True)

c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
c.login()

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

SamplesList = pd.read_csv('Z:\\SUPPLY CHAIN\\Python Scripts\\06_Samples\\SamplesList.csv', header = 0)

AllSamples = pd.merge(SamplesPlan, samples, how = 'outer', on = 'ConfigSKU', sort  = False)
AllSamples = pd.merge(AllSamples, SamplesList, how = 'outer', on = 'ConfigSKU', sort = False)
AllSamples = AllSamples.sort(columns = ['01_WHSamples_OUT'], ascending = True, na_position = 'last')
AllSamples = AllSamples[['Month','ConfigSKU','SKU','ConfigName','Supplier','Designer','Category','GLDay','GLMonth','GLYear', 'Buyer','Date QCed','01_WHSamples_OUT','02_SampleRoom_IN','03_SampleRoom_TO_studio','04_SampleRoom_FROM_studio','05_SampleRoom_OUT','06_WHSamples_IN']]

SampleSummary = AllSamples.groupby(['GLYear','GLMonth']).agg({'Month':'count','Date QCed':'count','01_WHSamples_OUT':'count','02_SampleRoom_IN':'count','03_SampleRoom_TO_studio':'count','04_SampleRoom_FROM_studio':'count','05_SampleRoom_OUT':'count','06_WHSamples_IN':'count'})
SampleSummary = SampleSummary[['Month','Date QCed','01_WHSamples_OUT','02_SampleRoom_IN','03_SampleRoom_TO_studio','04_SampleRoom_FROM_studio','05_SampleRoom_OUT','06_WHSamples_IN']]

OutputName = 'SampleTrack ' + str(today)
doc_name = 'Samples Tracker '
part = OutputName + '.xlsx'
message = 'Spree Samples Tracking ' + str(date.today())
maillist = "MailList.txt" 
     
writer = ExcelWriter(part)
SampleSummary.to_excel(writer,'EasyTrack')
AllSamples.to_excel(writer,'Master', index = False)
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
    
OutputName2 = 'SampleTrack'
part2 = OutputName2 + '.xlsx'
     
writer2 = ExcelWriter(part2)
AllSamples.to_excel(writer2,'Master', index = False)
workbook2 = writer2.book
wksht2 = writer2.sheets['Master']
wksht2.set_column('A:B', 10)
wksht2.set_column('C:C', 14)
wksht2.set_column('D:G', 18)
wksht2.set_column('H:K', 10)
wksht2.set_column('L:L', 18)
wksht2.set_column('M:R', 28)
writer2.save()
