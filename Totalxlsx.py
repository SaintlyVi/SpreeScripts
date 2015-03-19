# -*- coding: utf-8 -*-
"""
Created on Fri Jan 09 10:41:26 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
import os
#from datetime import date
from pandas import DataFrame, ExcelWriter

def data_total( DocName, HistoryPath, SavePath ):
    
    files = os.listdir(HistoryPath)
    df = DataFrame(data = files, columns = ['Files'])
    #df = df[df['Files'].str.contains('~') == False]

    TotalData = pd.DataFrame()
    
    for file in files:    
        historyfile = os.path.join(HistoryPath, file)
        try:
            HistoryBook = pd.ExcelFile(historyfile)
            HistorySheet = HistoryBook.parse('Sheet1', skiprows = 0, index = None)
            
            TotalData = TotalData.append(HistorySheet)
        
        except IOError:
            print "Cannot read " + str(historyfile)
    
    TotalData.dropna(subset = ['ProductID'], inplace = True)    
    
    filename = DocName + '.xlsx'
    filename = os.path.join(SavePath, filename)    
    
    writer = ExcelWriter(filename)
    TotalData.to_excel(writer, 'Sheet1', index = False )   
    writer.save()