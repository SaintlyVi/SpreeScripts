# -*- coding: utf-8 -*-
"""
Created on Thu Apr 23 10:29:41 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
from pandas import ExcelWriter
from datetime import date, timedelta
from openpyxl.reader.excel import load_workbook
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

today = date.today()
lastmonth = 1
nextmonth = 4

Visibility = AllData.InboundData(lastmonth, nextmonth, today)

Visibility['GLiso'] = Visibility['GLDate'].apply(lambda x: list(x.isocalendar())[:-1] + [3])
Visibility['GLgreg'] = Visibility['GLiso'].apply(lambda x: iso_to_gregorian(x[0],x[1],x[2]))

VConf = Visibility.drop_duplicates(subset = ['ConfigSKU','POs'])
GLTrack = VConf.groupby(['GLgreg'])