#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 11 11:54:13 2014

@author: Wiebke.Toussaint
"""
#This script collates all stock processed and counted by Spree QC teams. 
#Script is run manually by scanning agents when they have filled one pallet

import sys
sys.path.append('Z:\\Python Scripts\\00_SharedFunctions')

import pandas as pd
from pandas import DataFrame
from datetime import date, datetime
import gspread
from pandas import ExcelWriter
import MyFunx

team = raw_input('What is your Team Letter (A, B, C, S (S = SAMPLES))? ')
shift = datetime.today().strftime('%Y-%m-%d %H:%M')

#Import Brightpearl Detail Report
columns = ["Order ID", "Ref", "Status", "Contact", "SKU", "Name", "Quantity"]
BP = pd.read_csv('Z:\\Python Scripts\\00_UPDATE\\BPdetail.csv', header = 0, usecols = columns)
    
#Lulu1 = pd.ExcelFile('Z:\\Python Scripts\\00_UPDATE\\Lulu1.xlsx')
#Planned = Lulu1.parse('Sheet4', skiprows = 0, index = None, parse_cols = (36,46))

#Connect to MSSQL to get Lulu data
Lulu =  MyFunx.sql_import("vw_ProcurementPipeline","ActualGoLiveDate")
Planned = Lulu[['SimpleSKU','ProcurementProductCategoryL3']]
Planned.drop_duplicates(inplace = True, take_last = True)

#Reading from StockCount google docs
c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
c.login()

#==============================================================================
# Details for SAMPLES processing
#==============================================================================
if team == 'S':
    sht = c.open('SamplesCount')
    worksheet = sht.worksheet('Team' + team)
       
    #Parameters for total samples count doc
    DocName = 'Samples Returned'
    HistoryPath = 'Z:\\Stock Count\\All samples'
    SavePath = 'Z:\\Python Scripts\\02_StockCount'
    
    #Email Parameters for Spree output file
    OutputName = ' Team ' + str(team) + ' ' + str(date.today()) + ' ' + str(datetime.today().hour) + 'h' + str(datetime.today().minute)
    doc_name = 'Samples Count ' + 'Team ' + str(team) + ' '
    part = 'Samples Count '+ OutputName + '.xlsx'
    message = 'Spree Samples Count for shift ' + OutputName
    maillist = "MailList.txt"
    
    #BP1 = BP[BP['Ref'].str.contains('sample|Sample|SAMPLE|samples|Samples') == True]         
    
#==============================================================================
#  Details for processing of ALL other stock                   
#==============================================================================
else:
    sht = c.open('StockCount' + team)
    worksheet = sht.worksheet('Team' + team)
         
    #Parameters for total stock count doc
    DocName = 'Rolling Stock'
    HistoryPath = 'Z:\\Stock Count\\All handovers'
    SavePath = 'Z:\\Python Scripts\\02_StockCount'         
    
    #Email Parameters for Spree output file
    OutputName = 'Stock Count Team ' + str(team) + ' ' + str(date.today()) + ' ' + str(datetime.today().hour) + 'h' + str(datetime.today().minute)     
    doc_name = 'Spree Stock Count ' + 'Team ' + str(team) + ' '
    part = OutputName + '.xlsx'
    message = 'Spree Stock Count for shift ' + OutputName
    maillist = "MailList.txt"
    
    #BP1 = BP[BP['Ref'].str.contains('sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud') == False]      
           
#==============================================================================
# Computing sample/stock count
#==============================================================================

if worksheet.cell(2,1).value == "":
    print "Sheet", team, "is empty"
else:
    info = worksheet.get_all_values()
    headers = info.pop(0)
    df = DataFrame(data = info, columns = ['ProductID'])
    l = len(df)
    SKU = df['ProductID'].value_counts()
    count = DataFrame(data = SKU)
    count.reset_index(level=0, inplace=True)
    count.columns = ['ProductID', 'Qty Counted']
    dup = [d.upper() for d in count['ProductID']]
    count['ProductID'] = dup
    count['Date'] = shift
    count['Date'] = pd.to_datetime(count['Date'],coerce=True)
    count['Team'] = team

        #BP2 = BP[BP['Ref'].isnull()==True]
    BPdetail = BP#BP1.append(BP2, ignore_index=True)
    
    #Merging information from Lulu, Brighpearl and StockCount
    Merge = pd.merge(count, BPdetail, left_on='ProductID', right_on='SKU', how = 'left', sort = False)
    Merge1 = pd.merge(Merge, Planned, left_on = 'ProductID', right_on = 'SimpleSKU', how = 'left', sort = False)    
    ShiftCount = Merge1.rename(columns = {'Order ID':'PO', 'Contact':'Supplier', 'Quantity':'Qty Ordered', 'Name':'Description','ProcurementProductCategoryL3':'Category'})
    ShiftCount['Qty Diff'] = ShiftCount['Qty Counted'] - ShiftCount['Qty Ordered']
    ShiftCount = ShiftCount[['Date','PO','Status','Ref','Supplier','Qty Ordered','Qty Counted','Qty Diff','ProductID','Description','Category','Team']]
    ShiftCount = ShiftCount.sort(['Supplier', 'ProductID'], ascending=[1,1]) 
    ShiftCount.drop_duplicates(inplace = True)   
    
    #Create Spree excel output file    
    writer = ExcelWriter(part)
    ShiftCount.to_excel(writer,'Sheet1', index = False)
    workbook = writer.book
    wksht = writer.sheets['Sheet1']
    wksht.set_column('A:A', 18)
    wksht.set_column('B:B', 8)
    wksht.set_column('C:E', 20)
    wksht.set_column('F:H', 10)
    wksht.set_column('I:I', 15)
    wksht.set_column('J:J', 45)
    wksht.set_column('K:K', 12)
    wksht.set_column('L:L', 6)
    writer.save()
    
    MyFunx.data_total(DocName, HistoryPath, SavePath )
    
    #Deleting all data from google doc    
    clean = worksheet.range('A2:A' + str(l + 5))
    for cl in clean:
        cl.value = ""
    worksheet.update_cells(clean)
    
    MyFunx.send_message(doc_name, message, part, maillist)
    