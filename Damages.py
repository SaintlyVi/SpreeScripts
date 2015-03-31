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

#Reading from QCDamages google doc
c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
c.login()

sht = c.open('QCDamages')
worksht = sht.worksheet('Sheet1')
if worksht.cell(2,1).value == "":
    print "No damages scanned"
    Damaged = DataFrame(columns = ['ProductID', 'Reason for damage', 'QC Responsible','Date', 'Damaged'])
    lgth = 0
else:
    info = worksht.get_all_values()
    headers = info.pop(0)
    dmgs = DataFrame(data = info, columns = ['ProductID', 'Reason for damage', 'QC Responsible'])
    lgth = len(dmgs)
    dup = [d.upper() for d in dmgs['ProductID']]
    dmgs['ProductID'] = dup
    dmgs['Date'] = today
    Damaged = dmgs.groupby(['ProductID', 'Reason for damage']).aggregate({'ProductID':'size'})
    Damaged.columns = ['Damaged']
    Damaged.reset_index(inplace = True)

#Reading from Oversupply google doc    
sheet = c.open('Oversupply')
ws = sheet.worksheet('Sheet1')
if ws.cell(2,1).value == "":
    print "No oversupply scanned"
    OS = DataFrame(columns = ['SKU','OS'])
    l = 0
else:
    info = ws.get_all_values()
    headers = info.pop(0)
    OS = DataFrame(data = info, columns = ['SKU'])
    l = len(OS)
    dup = [d.upper() for d in OS['SKU']]
    OS['SKU'] = dup
    OS = OS.groupby('SKU').agg({'SKU':'size'})
    OS.columns = ['OS']
    OS.reset_index(inplace = True)
        
#Import Brightpearl Detail Report
columns = ["Order ID", "Contact", "SKU", "Name"]
BPdetail = pd.read_csv('BPdetail.csv', header = 0, usecols = columns, dtype = {'Order ID': unicode})

#Import Supplier Contacts
Contacts = pd.ExcelFile('03_Damages\\Supplier Contacts.xlsx')
Contacts = Contacts.parse('Sheet1', skiprows = 0, index = None, parse_cols = (1,2,5))
Contacts.drop_duplicates(subset = ['POs'], inplace = True, take_last = True)
 
#Merge Brightpearl and Damages data
Merge = pd.merge(Damaged, OS, left_on='ProductID', right_on='SKU', how = 'outer')
Merge.loc[Merge.SKU.isnull(),"SKU"] = Merge.ProductID
Merge = pd.merge(Merge, BPdetail, on='SKU', how = 'left')

Merge = pd.merge(Merge, Contacts, left_on = 'Order ID', right_on = 'POs', how = 'left')
Merge = Merge[['Contact','Order ID','SKU','Name','Reason for damage','Damaged','OS','Client name', 'Client email']]

DayCount = Merge.rename(columns = {'Order ID':'PO', 'Contact':'Supplier', 'Name':'Description'})
DayCount = DayCount.sort(['Supplier','PO','SKU'], axis=0, ascending=[1,1,1])
DayCount['Date'] = today
cols = DayCount.columns.tolist() #rearrange columns
cols = cols[-1:] + cols[:-1]
DayCount = DayCount[cols]

#Create Spree excel output file
writer = ExcelWriter('03_Damages\\Damages ' + str(today) + '.xlsx')
DayCount.to_excel(writer,'Sheet1', index = False)

#Format excel doc
workbook = writer.book
wksht = writer.sheets['Sheet1']
wksht.set_column('A:A', 8)
wksht.set_column('B:B', 25)
wksht.set_column('C:C', 12)
wksht.set_column('D:D', 18)
wksht.set_column('E:E', 35)
wksht.set_column('F:F', 20)
wksht.set_column('G:G', 8)
wksht.set_column('H:H', 8)
wksht.set_column('I:J', 35)
writer.save()

#Deleting all data from google doc
clean1 = worksht.range('A2:C' + str(lgth + 5))
for cl in clean1:
    cl.value = ""
worksht.update_cells(clean1)
    
clean2 = ws.range('A2:C' + str(l + 5))
for cl in clean2:
    cl.value = ""
ws.update_cells(clean2)

#Email Spree output file 
doc_name = 'Daily Damages ' 
part = '03_Damages\\Damages ' + str(today) + '.xlsx'
message = 'Daily Damages'
maillist = 'MailList_Damages.txt'
MyFunx.send_message(doc_name, message, part, maillist)

#Create 8 week rolling damages doc
path = '03_Damages'    
DataName = dmgs
DocName = 'Rolling Damages'
DaysCounting = 56
MyFunx.data_history( DataName, DocName, DaysCounting, path )

#Create 8 week rolling oversupply doc
Oversup = DayCount[["Date","Supplier","PO","SKU","OS"]]
path = '03_Damages'    
DataName = Oversup
sheet = 'Sheet1'
DocName = 'Rolling Oversupply'
DaysCounting = 56
MyFunx.data_history( DataName, DocName, DaysCounting, path, sheet )


