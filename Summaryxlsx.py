# -*- coding: utf-8 -*-
"""
Created on Fri Jan 09 14:24:33 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
import os
#from datetime import date
from pandas import DataFrame, ExcelWriter

def data_summary( DocName, SavePath, HistoryPath, sheet ):
    
    files = os.listdir(HistoryPath)
    df = DataFrame(data = files, columns = ['Files'])
    df = df[~df['Files'].str.contains('~')]

    SummaryData = pd.DataFrame()
    
    for file in files:    
        historyfile = os.path.join(HistoryPath, file)
        try:
            HistoryBook = pd.ExcelFile(historyfile)
            HistorySheet = HistoryBook.parse( sheet, skiprows = 0, index = None)
            
            SummaryData = SummaryData.append(HistorySheet)
        
        except IOError:
            continue
    
    SummaryData.drop_duplicates(inplace = True)    
    
    filename = DocName + '.xlsx'
    filename = os.path.join(SavePath, filename)    
    
    writer = ExcelWriter(filename)
    SummaryData.to_excel(writer, 'Sheet1', index = False )   
    writer.save()