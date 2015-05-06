# -*- coding: utf-8 -*-
"""
Created on Wed Apr 08 12:12:49 2015

@author: Wiebke.Toussaint
"""

#==============================================================================
##VISIBILITY
##Product Tracking Report looking 5 weeks back, 1 week forward
##TO DO before running the script:
##1. Download -5week + 1 week Brightpearl Detail Report (filter DELIVERY DATE)
##2. Download -2month + 1 month Brightpearl PO Report
##3. Refresh Lulu PowerQuery
##4. Refresh IBOI1003 PowerQuery
##5. Run Damages Script
##6. Run Stock Count Script
#==============================================================================
def InboundData(lastmonth, nextmonth, today):
    
    import numpy as np
    import pandas as pd
    from pandas import DataFrame
    import gspread
    import MyFunx    
    
    #==============================================================================
    # Import from all required data sources
    #==============================================================================
    #Import Brightpearl Detail Report data
    columns = [u"Order ID", u"Ref", u"SKU", u"Status", u"Quantity"]
    BPdetail = pd.read_csv('BPdetail.csv', header = 0, usecols = columns, encoding = 'iso-8859-1')
    BPdetail['Order ID'] = BPdetail['Order ID'].map(lambda x: unicode(x))
    BPdetail.rename(columns={'Order ID': u'POs', 'Quantity': u'BP Qty'}, inplace=True)
    BPdet = BPdetail[BPdetail['Status'].str.contains('Cancel PO')==False]
    
    BPdet = pd.pivot_table(BPdet, values = [u'BP Qty'], index = [u'SKU',u'POs',u'Ref',u'Status'], aggfunc=np.sum)
    BPdet.reset_index(inplace=True)
       
    CancelledPOs = BPdetail[BPdetail['Status']=='Cancel PO']
    CancelledPOs = CancelledPOs.groupby('POs').agg({'SKU':'count'})
    
    #Import Brightpearl PO Report data
    columns = ["Order ID", "Delivery due"]
    BPreport = pd.read_csv('BPreport.csv', header = 0, usecols = columns, parse_dates = [1])
    BPreport = BPreport.dropna(axis = 0,how = 'all') #removes empty rows
    BPreport['Order ID'] = BPreport['Order ID'].map(lambda x: x.strip('PO#')) #removes text in front of PO number
    BPreport.rename(columns={'Order ID': 'POs', 'Delivery due':'DeliveryDue'}, inplace=True)
    
    BP = pd.merge(BPdet, BPreport, on = 'POs', how = 'left', sort = False)
    
    #Import Epping Receiving Report data
    c = gspread.Client(auth=('spreewarehouse@gmail.com', 'spreeapp'))
    c.login()
    
    sht = c.open('Epping Receiving Report')
    worksheet = sht.worksheet('Booked')
    
    info = worksheet.get_all_values()
    headers = info.pop(0)
    B_R = DataFrame(data = info, columns = headers)
    
    Bookd = B_R[[u'POs',u'Date booked']]
    Bookd = Bookd.replace('',np.nan)
    Bookd = Bookd.dropna(subset = [u'Date booked'], thresh = 1)
    Bookd = Bookd.drop_duplicates(subset = [u'POs'], take_last = False)
    Bookd = Bookd.loc[Bookd.POs != ""]
    Bookd[u'Date booked'] = pd.to_datetime(Bookd[u'Date booked'], infer_datetime_format = True)
    
    Receivd = B_R[['POs', 'Partial delivery', 'Date received']]
    Receivd = Receivd.replace('',np.nan)
    Receivd = Receivd.dropna(subset = ['Date received'], thresh = 1)
    Receivd = Receivd.drop_duplicates(subset = ['POs'], take_last = True)
    Receivd = Receivd.loc[Receivd.POs != ""]
    Receivd['Date received'] = pd.to_datetime(Receivd['Date received'], infer_datetime_format = True)
    
    #Import Rolling Stock data
    Stock = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\02_StockCount\\Rolling Stock.xlsx')
    QCed = Stock.parse('Sheet1', skiprows = 0, index = None, parse_cols = (1,3,4,5), encoding = 'utf-8')
    QCed.rename(columns={'Date': u'LastQCed', 'PO':u'POs','ProductID':u'SKU'}, inplace=True)
    poqc = [unicode(p) for p in QCed['POs']]
    QCed['POs'] = poqc
    QCed = QCed.groupby(['POs','SKU']).agg({'Qty Counted':np.sum, 'LastQCed':np.max})
    QCed.reset_index(inplace=True)
    
    #Import Rolling Damages
    Damages = pd.ExcelFile('03_Damages_OS\\Rolling Damages.xlsx')
    Damages = Damages.parse('Sheet1', skiprows = 0, index = None, encoding = 'utf-8')
    SKU = Damages['SKU'].value_counts()
    Damagd = DataFrame(data = SKU)
    Damagd.reset_index(level=0, inplace=True)
    Damagd.columns = [u'SKU', u'Qty Damaged']
    
    #Import Lulu Assortment Plans
    ## table = "vw_ProcurementPipeline"
    ## dateparse = "ActualGoLiveDate"
    
    pw = raw_input("Enter SQL Server database password: ")
    Lulu =  MyFunx.sql_import("vw_ProcurementPipeline",["ActualGoLiveDate","PlannedGoLiveDate"], pw)
    Planned = Lulu[['PlannedGoLiveDayOfWeek','PlannedGoLiveMonth','PlannedGoLiveYear','PlannedGoLiveDate','BuyerPlanName','BuyerPlanStatus','EmployeeFirstName','PlannedUnitCostExclTax','PlannedTotalQuantity','PlannedTotalCostExclTax','SimpleSKU','SimpleName','ConfigName','ConfigSKU','ProcurementStatus','ProcurementProductCategoryL3','ActualGoLiveDate','Supplier','Designer','EANNumber','BarCode']]
    #Merge EAN, BarCode information with SKU
    Planned.loc[Planned.EANNumber == "",'EANNumber'] = None
    SKU = Planned['EANNumber'].combine_first(Planned['SimpleSKU'])
    Planned.loc[:,'SKU'] = Planned['BarCode'].combine_first(SKU)
    Planned.drop_duplicates(subset = ['SKU','PlannedGoLiveMonth'], inplace = True, take_last = True)
    if lastmonth == 11 | 12:
        Planned = Planned[((Planned.PlannedGoLiveMonth >= lastmonth) & (Planned.PlannedGoLiveYear == today.year)) | ((Planned.PlannedGoLiveMonth <= nextmonth) & (Planned.PlannedGoLiveYear == today.year))]
    else:    
        Planned = Planned[((Planned.PlannedGoLiveMonth >= lastmonth) & (Planned.PlannedGoLiveYear == today.year)) & ((Planned.PlannedGoLiveMonth <= nextmonth) & (Planned.PlannedGoLiveYear == today.year))]
    Planned.rename(columns={'PlannedGoLiveDayOfWeek':'GLDay','PlannedGoLiveMonth':'GLMonth','PlannedGoLiveYear':'GLYear','PlannedGoLiveDate':'GLDate', 'EmployeeFirstName':'Buyer','ProcurementProductCategoryL3':'Category', 'PlannedUnitCostExclTax':'UnitCost','PlannedTotalQuantity':'TotalUnits','PlannedTotalCostExclTax':'TotalCost'}, inplace=True)
    Planned = Planned[Planned['TotalCost'] > 0]
    Planned.loc[Planned.ProcurementStatus == 'Deleted', ['TotalUnits','TotalCost']] = ""
    
    #Import Dynaman IBOI1003 Inbound Order Received Messages
    ## table = "vw_WarehouseInboundItemsReceived"
    ## dateparse = "Timestamp"
    
    IBOI1003 =  MyFunx.sql_import("vw_WarehouseInboundItemsReceived","Timestamp",pw)
    IBOI1003 = IBOI1003[['MessageReference','ItemCode','QuantityReceived','Timestamp']]
    TaknIn = IBOI1003.groupby(['MessageReference','ItemCode']).agg({'QuantityReceived':np.sum, 'Timestamp':np.max})
    TaknIn.reset_index(inplace=True)
    TaknIn.columns = ['POs','SKU','OTDLastReceived','Qty Received']
    TaknIn['POs'] = TaknIn['POs'].apply(lambda x: x if len(x) < 7 else 0)
    TaknIn = TaknIn[TaknIn['POs'] != 0]
    
    #Import Dynaman ITMI1002 
    ## table = "vw_WarehouseStockAvailability"
    ## dateparse = "Timestamp"
    
    ITMI1002 =  MyFunx.sql_import("vw_WarehouseStockAvailability","Timestamp",pw)
    ITMI1002 = ITMI1002[['ITEM_CODE','QTY']]
    PutAway = pd.pivot_table(ITMI1002, values = ['QTY'], index = ['ITEM_CODE'], aggfunc=np.sum)
    PutAway.reset_index(inplace=True)
    PutAway.columns = ['SKU','Qty PutAway']
    
    Merge = pd.merge(Planned, BP, on = 'SKU', how = 'left', sort = False)
    Merge1 = pd.merge(Merge, Bookd, on = 'POs', how = 'left', sort = False)
    Merge2 = pd.merge(Merge1, Receivd, on = 'POs', how = 'left', sort = False)
    Merge3 = pd.merge(Merge2, QCed, on = ['SKU','POs'], how = 'left', sort = False)
    Merge4 = pd.merge(Merge3, Damagd, on = 'SKU', how = 'left', sort = False)
    Merge5 = pd.merge(Merge4, TaknIn, on = ['SKU','POs'], how = 'left', sort = False)
    Visibility = pd.merge(Merge5, PutAway, on = 'SKU', how = 'left')
    Visibility.drop_duplicates(inplace = True)
    #remove duplicates that are in Draft PO status,assuming that these POs are outdated
    Visibility.loc[Visibility['Ref'].str.contains("sample|Sample|SAMPLE|samples|Samples|OS|Os|OVERSUPPLY|fraud")==True,['TotalUnits','TotalCost','Qty PutAway']] = ""
    Visibility['Oversupply'] = ""
    Visibility.loc[Visibility['Ref'].str.contains("OS|Os|OVERSUPPLY")==True, 'Oversupply'] = Visibility['BP Qty']    
    V0 = Visibility.loc[Visibility.TotalUnits == "",]
    Vn0 = Visibility.loc[Visibility.TotalUnits != "",]
    Va = Vn0[Vn0.duplicated(subset = ['SKU'], take_last = False)==False]
    Vb = Vn0[Vn0.duplicated(subset = ['SKU'], take_last = False)==True]
    Vb.loc[:, ['TotalUnits','TotalCost','Qty PutAway']] = ""    
    Vc = Vb[Vb['Status']!='Draft PO'] 
    Visibility = Va.append([Vc,V0], ignore_index = True)
    Visibility.replace("", np.nan, inplace = True)
    
    return Visibility