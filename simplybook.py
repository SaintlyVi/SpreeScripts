# -*- coding: utf-8 -*-
"""
Created on Wed Jun 10 12:40:46 2015

@author: Wiebke.Toussaint
"""

import requests
import json
import pandas as pd

#Create access token

def getToken():
    API_KEY = '3b52bf40044d75d52b257a1c41a2706947d3c4fcf561a258028a63b7f8d284d4'
    COMPANY_LOGIN = 'Spree'
    url = 'http://user-api.simplybook.me/login'
    headers = {
    'Content-Type': 'application/json; charset=UTF-8',
    'Content-Length': '130',
    'Accept': 'application/json'}
    r = requests.post(url, json = {"jsonrpc":"2.0","method":"getToken","params":[COMPANY_LOGIN,API_KEY],"id":1}, headers = headers)
    return(r.json()['result'])
    
#Create headers
def createHeaders():
    headers = {
    'Content-Length': '60',
    'Content-Type': 'application/json; charset=UTF-8',
    'X-Company-Login': 'Spree',
    'X-Token': getToken(),
    'Accept': 'application/json'}
    return headers

#Fetch data from simplybook
url = 'https://user-api.simplybook.me/'
getReservedTime = requests.post(url, json = {"jsonrpc":"2.0","method":"getReservedTime","params":['2015-06-01','2015-06-10','1','1'],"id":1}, headers = createHeaders())

ReservedTime = pd.read_json(getReservedTime.content).result.reset_index()
    
#Request-Line
    #[Method]POST [Request-URI:abs_path]/login [protocol version]HTTP/1.1
#[URI(authority: Host header field]Host: user-api.simplybook.me
#[entity header 1]Content-Type: application/json; charset=UTF-8
#[entity header 2...signals presence of message]Content-Length: 130
#[header field 1]Accept: application/json

#Message Body
    #{"jsonrpc":"2.0","method":"getToken","params":["spree","3b52bf40044d75d52b257a1c41a2706947d3c4fcf561a258028a63b7f8d284d4"],"id":1}