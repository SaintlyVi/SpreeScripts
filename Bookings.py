#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 11 11:54:13 2014

@author: Wiebke.Toussaint
"""

## This script takes bookings from Simplybook, createst the next day's Receiving Report and emails
## it to buyers and warehouse. All suppliers' contact information is recorded in a 6 week rolling 
## doc

import pandas as pd
from pandas import DataFrame, Series
from datetime import date, timedelta
from pandas import ExcelWriter
import MyFunx

today = date.today()

#Retrieve data from Simplybook export
columns = ["Date","Time","Event","Client name","Client email","Is cancelled","Record date","Additional fields"]
SBexport = pd.read_csv('SBexport.csv', skiprows = [1], header = 1, usecols = columns, parse_dates=[0], dayfirst=True)

#Separate all additional fields
a = SBexport["Additional fields"].str.split(';').apply(Series, 1).unstack()
s = a.str.split(': ').apply(Series, 1)
brand = s[1][0]
pos = s[1][1]
boxes = s[1][2]
partial_delivery = s[1][3]
supplier = s[1][4]

#Merge additional fields with other booking data
idx = ['Supplier', 'Brand', 'POs', '# boxes', 'Partial delivery']
adf = DataFrame(data = [supplier, brand, pos, boxes, partial_delivery], index = idx).T
df = pd.merge(SBexport, adf, left_index=True, right_index=True)

#Split entries with multiple POs
x = df['POs'].str.split('\n|/ |, | ').apply(Series, 1).stack()
x = x.map(lambda y: y.strip(" ()#abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-&:.%"))
x.index = x.index.droplevel(-1)
x.name = 'POs'
del df['POs']
m = df.join(x)

#Create Supplier Contacts output files
S = m[['Date','Event','POs','Supplier','Brand','Client name','Client email']]
S = S[S['POs'].isnull() == False]
path = '03_Damages_OS'
DataName = S
DocName = 'Supplier Contacts'
DaysCounting = 70
MyFunx.data_history( DataName, DocName, DaysCounting, path )

#Select tomorrow's date    
dates = []
for d in m['Date']:
    if today.weekday() == 4:
        if d.date() == today + timedelta(3):
            d = d.strftime('%Y/%m/%d')
    elif d.date() == today + timedelta(1):
            d = d.strftime('%Y/%m/%d')
    else: 
        d = 0
    dates.append(d)
m["Date"] = dates

#Clean up date and time values
time = [r.split('-')[0] for r in m["Time"]]
m["Time"] = time 

date_booked = [r.split(' ')[0] for r in m["Record date"]]
m["Record date"] = date_booked

#Create Booking output file
R = m[['Date', 'Time', 'Event', 'POs', 'Supplier', 'Brand', '# boxes', 'Partial delivery']]
R2 = R[R['POs'] != ""] #delete rows with no PO information
R3 = R2[R2['Date'] != 0] #delete rows with date after today
RowCount = len(R3.index)
Receiving = R3.set_index([range(0, RowCount, 1)])
Receiving = Receiving.drop_duplicates(subset = ['POs'])

writer = ExcelWriter('01_Bookings\\Bookings ' + str(today + timedelta(1)) + '.xlsx')
Receiving.to_excel(writer,'Sheet1', index = False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']
format = workbook.add_format({'text_wrap': True})
worksheet.set_column('A:B', 10 )
worksheet.set_column('C:C', 35 )
worksheet.set_column('D:D', 10 )
worksheet.set_column('E:F', 35 )
worksheet.set_column('G:H', 14, format )
writer.save()

#Email Spree output file
doc_name = 'Bookings ' 
part = '01_Bookings\\Bookings ' + str(today + timedelta(1)) + '.xlsx'
message = 'Supplier Bookings for ' + str(today + timedelta(1))
maillist = 'MailList_Bookings.txt'
MyFunx.send_message(doc_name, message, part, maillist)