#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Fri Nov 21 14:40:29 2014

@author: Wiebke.Toussaint
"""

#This script reads all damaged stock processed by Spree QC teams. 
#Script must be scheduled to run on a daily basis at 23:45.

import pandas as pd
from pandas import DataFrame
from datetime import date
import gspread
from pandas import ExcelWriter
import MyFunx

today = date.today()

#Reading from QCDamages google docs
c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
c.login()

sht = c.open('QCDamages')
worksheet = sht.worksheet('Sheet1')
if worksheet.cell(2,1).value == "":
    print "No damages scanned"
else:
    info = worksheet.get_all_values()
    headers = info.pop(0)
    dmgs = DataFrame(data = info, columns = ['ProductID', 'Reason for damage', 'QC Responsible'])
    l = len(dmgs)
    dup = [d.upper() for d in dmgs['ProductID']]
    dmgs['ProductID'] = dup
    dmgs['Date'] = [today]*l
    dmgs['Date'] = pd.to_datetime(dmgs['Date'], coerce=True)
    
sheet = c.open('Oversupply')
worksheet = sheet.worksheet('Sheet1')
if worksheet.cell(2,1).value == "":
    print "No oversupply scanned"
else:
    info = worksheet.get_all_values()
    headers = info.pop(0)
    OS = DataFrame(data = info, columns = ['SKU'])
    l = len(OS)
    dup = [d.upper() for d in OS['SKU']]
    OS['SKU'] = dup
        
#Import Brightpearl Detail Report
columns = ["Order ID", "Contact", "SKU", "Name"]
BPdetail = pd.read_csv('BPdetail.csv', header = 0, usecols = columns, dtype = {'Order ID': unicode})

#Import Supplier Contacts
Contacts = pd.ExcelFile('03_Damages\\Supplier Contacts.xlsx')
Contacts = Contacts.parse('Sheet1', skiprows = 0, index = None, parse_cols = (1,2,5))    

#Merge Brightpearl and Damages data
Merge = pd.merge(dmgs, BPdetail, left_on='ProductID', right_on='SKU', how = 'right')
Merge = pd.merge(Merge, Contacts, left_on = 'ProductID', right_on = 'POs', how = 'left')    
Merge = Merge[['Date','Contact','Order ID','ProductID','Name','Reason for damage','QC Responsible', 'Client name', 'Client email']]
DayCount = Merge.rename(columns = {'Order ID':'PO', 'ProductID':'SKU', 'Contact':'Supplier', 'Name':'Description'})
DayCount = DayCount.sort(['Supplier','PO','SKU'], axis=0, ascending=[1,1,1])

#Create Spree excel output file
writer = ExcelWriter('03_Damages\\Damages ' + str(today) + '.xlsx')
DayCount.to_excel(writer,'Sheet1', index = False)

#Format excel doc
workbook = writer.book
wksht = writer.sheets['Sheet1']
wksht.set_column('A:A', 22)
wksht.set_column('B:B', 35)
wksht.set_column('C:C', 10)
wksht.set_column('D:D', 15)
wksht.set_column('E:E', 50)
wksht.set_column('F:F', 35)
wksht.set_column('G:G', 15)
wksht.set_column('H:H', 35)
wksht.set_column('I:I', 50)
writer.save()

#Deleting all data from google doc    
clean = worksheet.range('A2:C' + str(l + 5))
for cl in clean:
    cl.value = ""
worksheet.update_cells(clean)

#Email Spree output file 
doc_name = 'Daily Damages ' 
part = '03_Damages\\Damages ' + str(today) + '.xlsx'
message = 'Daily Damages'
maillist = 'MailList_Damages.txt'
MyFunx.send_message(doc_name, message, part, maillist)

#Create 6 week rolling doc
path = '03_Damages'    
DataName = DayCount
DocName = 'Rolling Damages'
DaysCounting = 56
MyFunx.data_history( DataName, DocName, DaysCounting, path )


