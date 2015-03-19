# -*- coding: utf-8 -*-
"""
Created on Wed Nov 26 12:27:06 2014

@author: Wiebke.Toussaint
"""

# This function creates a rolling history summary of data and  writes it to a xlsx file.
# INPUT PARAMETERS:
# DataName = Name of DataFrame in script that is rolled up
# DocName = Name of .xlsx file to which data will be written (must exist and have
# headers that corresond to DataFrame column.values. Date value must be named 'Date')
# DaysCounting = number of days of data to keep

import pandas as pd
import os
from datetime import date, timedelta
from pandas import ExcelWriter

def data_history( DataName, DocName, DaysCounting, path ):
    
    filename = DocName + '.xlsx'
    filename = os.path.join(path, filename)
    
    HistoryBook = pd.ExcelFile(filename)
    HistorySheet = HistoryBook.parse('Sheet1', skiprows = 0, index = None)    
    NewData = HistorySheet.append(DataName)
    AllDates = []
    for d in NewData['Date']:   
        if d.date() < date.today() - timedelta(DaysCounting):
            d = 0
        else:
            d = d
        AllDates.append(d)
    NewData['Date'] = AllDates
                
    NewData = NewData[NewData['Date']!= 0]
     
    writer = ExcelWriter(filename)
    NewData.to_excel(writer, 'Sheet1', index = False )   
    writer.save()

          
   