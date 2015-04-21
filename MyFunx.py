# -*- coding: utf-8 -*-
"""
Created on Thu Mar 19 14:08:32 2015

A collection of functions to be applied in various Spree Scripts.

@author: Wiebke.Toussaint
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import Encoders
import pandas as pd
import os
from datetime import date, timedelta
from pandas import ExcelWriter
import pyodbc
import pandas.io.sql as psql

def send_message( doc_name, message, part, maillist ): 
    ## This function sends an email + attachment.

    msg = MIMEMultipart()
            
    today = date.today()
    urlFile = open(maillist, "r+") 
    maillist = [i.strip() for i in urlFile.readlines()]    
        
    fromEmail = 'spreewarehouse@gmail.com' 
    #create message
    msg['Subject'] = str(doc_name) + str(today)
    msg['From'] = fromEmail
    #msg['To'] = ', '.join(MailList)
    body = message
    content = MIMEText(body, 'plain')
    msg.attach(content)
        
    #create attachment        
    filename = str(part)
    f = file(filename)
    attachment = MIMEText(f.read())
    attachment.set_payload(open(part, 'rb').read())
    Encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)           
    msg.attach(attachment)
        
    #call server and send email      
    mailServer = smtplib.SMTP('smtp.gmail.com', 587)
    mailServer.set_debuglevel(1)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    password = "spreeapp"
    mailServer.login(fromEmail, password)
    mailServer.ehlo()
    mailServer.sendmail(fromEmail, maillist, msg.as_string())
    mailServer.quit()
        
    print "Mail sent successfully"
       
       
def data_history( DataName, DocName, DaysCounting, path, sheet = 'Sheet1' ):
    # This function creates a rolling history summary of data and  writes it to a xlsx file.
    # INPUT PARAMETERS:
    # DataName = Name of DataFrame in script that is rolled up
    # DocName = Name of .xlsx file to which data will be written (must exist and have
    # headers that corresond to DataFrame column.values. Date value must be named 'Date')
    # DaysCounting = number of days of data to keep
    
    filename = DocName + '.xlsx'
    filename = os.path.join(path, filename)
    
    HistoryBook = pd.ExcelFile(filename)
    HistorySheet = HistoryBook.parse(sheet, skiprows = 0, index = None)    
    NewData = HistorySheet.append(DataName)
    period = date.today() - timedelta(days = DaysCounting)
    NewData['Age'] = NewData.loc[:,'Date'] - period
    NewData.reset_index(drop = True, level = 0, inplace = True)
    NewData = NewData[NewData['Age'] >= 0]
                    
    writer = ExcelWriter(filename)
    NewData.to_excel(writer, sheet, index = False)   
    writer.save()

def sql_import(table, dateparse, pw):
    # This function pulls data from SQL server    
    
    connection_string = "DRIVER={SQL Server};SERVER=02CPT-TLSQL01;DATABASE=Spree SSBI;UID=SSBI_PaymentOps;PWD=" + pw + ";TABLE=" + table
    cnxn = pyodbc.connect(connection_string, charset = 'utf8', unicode_results=True)
    sql = "select * from " + table
    df = psql.read_sql(sql, cnxn, parse_dates = dateparse)
    return df
    

def data_total( DocName, HistoryPath, SavePath ):
    
    files = os.listdir(HistoryPath)
    
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
    TotalData.drop_duplicates(inplace = True)    
    
    filename = DocName + '.xlsx'
    filename = os.path.join(SavePath, filename)    
    
    writer = ExcelWriter(filename)
    TotalData.to_excel(writer, 'Sheet1', index = False )   
    writer.save()