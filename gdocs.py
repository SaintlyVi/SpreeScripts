# -*- coding: utf-8 -*-
"""
Created on Wed May 27 16:05:32 2015

@author: Wiebke.Toussaint
"""

import requests, gspread, os, ast
from oauth2client.client import SignedJwtAssertionCredentials

def authenticate_gdocs():
    f = file(os.path.join('SpreeStockCount-95b566801cc4.p12'), 'rb')
    SIGNED_KEY = f.read()
    f.close()
    scope = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
    credentials = SignedJwtAssertionCredentials('username@gmail.com', SIGNED_KEY, scope)

    data = {
        'refresh_token' : '1/gqqYJ7LKCR5xHlt0ADlhhbYQYJUi-w41HPA5_FOOtCgMEudVrK5jSpoR30zcRFq6',
        'client_id' : '293040254780-ovm9vrh8cl18fib06vb5er913s09ppg7.apps.googleusercontent.com',
        'client_secret' : 'AdGjPdVC6pBp7DJHhv45Pvkr',
        'grant_type' : 'refresh_token',
    }

    r = requests.post('https://accounts.google.com/o/oauth2/token', data = data)
    credentials.access_token = ast.literal_eval(r.text)['access_token']

    c = gspread.authorize(credentials)
    return c