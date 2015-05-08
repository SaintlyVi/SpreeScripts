# -*- coding: utf-8 -*-
"""
Created on Tue Apr 21 12:47:00 2015

@author: Wiebke.Toussaint
"""

import pandas as pd
import numpy as np

In = pd.ExcelFile('Z:\\SUPPLY CHAIN\\Python Scripts\\05_TransportCost\\TcostFile.xlsx')
Orders = In.parse('Sheet1', skiprows = 0, index = None, parse_dates = True)
Orders['Type'] = Orders['Shipment ID'].apply(lambda x : x[:3])
Returns = Orders.loc[(Orders['Type'] == 'SPR') | (Orders['Type'] == 'COL'),:] #select return orders
Returns['Delay'] = Returns.apply(lambda x : x.PODDate - x.POHDate if x['PODDate'] != None else np.nan, axis = 1) #calculate delay from return order received on Dynaman to courrier pickup
Returns.loc[Returns.Delay < 0, 'Delay'] = np.nan #set delay of returns that only get logged at WH to zero
CourierReturns = Returns.loc[Returns['Charge Out'] > 0] #remove multiple parcels on same order from total count
print CourierReturns[['Charge Out','Delay']].describe()
print 'Total Charge Out Courier Collections ' + str(CourierReturns['Charge Out'].sum())

#Post Office Returns
Returns['PO Charge'] = np.nan
Returns.loc[Returns['Chargeable Mass'] <= 1000, 'PO Charge'] = 40.90 + 6.75
Returns.loc[Returns['Chargeable Mass'] > 1000, 'PO Charge'] = 40.90 + 5.50*np.ceil(Returns['Chargeable Mass']/1000.0) + 6.75
print Returns[['PO Charge']].describe()
print 'Total equivalent charge out Post Office drop-offs ' + str(Returns['PO Charge'].sum())

#PostNet Returns
Returns['PNet Charge'] = np.nan
#Main Centre under 2kg
Returns.loc[(Returns['Chargeable Mass'] <= 2000) & (Returns['Zone'].isin(['Main Centre','Local','Main Township','Local Township'])), 'PNet Charge'] = 54.57
#Regional under 2kg
Returns.loc[(Returns['Chargeable Mass'] <= 2000) & (~Returns['Zone'].isin(['Main Centre','Local','Main Township','Local Township'])), 'PNet Charge'] = 81.43
#Main Centre over 2kg
Returns.loc[(Returns['Chargeable Mass'] > 2000) & (Returns['Zone'].isin(['Main Centre','Local','Main Township','Local Township'])), 'PNet Charge'] = 54.57 + 4.95*np.ceil(Returns['Chargeable Mass']/1000.0)
#Regional over 2kg
Returns.loc[(Returns['Chargeable Mass'] > 2000) & (~Returns['Zone'].isin(['Main Centre','Local','Main Township','Local Township'])), 'PNet Charge'] = 81.43 + 6.86*np.ceil(Returns['Chargeable Mass']/1000.0)

print Returns[['PNet Charge']].describe()
print 'Total equivalent charge out PostNet drop-offs ' + str(Returns['PNet Charge'].sum())

