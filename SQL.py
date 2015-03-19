# -*- coding: utf-8 -*-
#!/usr/bin/python
"""
Created on Wed Oct 29 15:15:51 2014

@author: Wiebke.Toussaint
"""

import pyodbc
import pandas.io.sql as psql

def sql_import(table, dateparse):
    connection_string = "DRIVER={SQL Server};SERVER=02CPT-TLSQL01;DATABASE=Spree SSBI;UID=SSBI_PaymentOps;PWD=Spr33Pops101;TABLE=" + table
    cnxn = pyodbc.connect(connection_string)
    sql = "select * from " + table
    df = psql.read_sql(sql, cnxn, parse_dates = [dateparse])
    return df