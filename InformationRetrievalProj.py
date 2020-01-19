# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 20:23:23 2020

@author: benic and mikeyg
"""
import pandas as pd
import os

from openpyxl import load_workbook
import nltk

def iter_rows(ws):
    result=[]
    for row in ws.iter_rows():
        rowlist = []
        for cell in row:
            rowlist.append(cell.value)
        result.append(rowlist)
    return result

workbook = load_workbook(filename='data.xlsx', read_only = True)
worksheet = workbook['sheet']
fileList =  (list(iter_rows(worksheet)))
geoLocation = []
tweet = []
i=2
for col in fileList:
    geoLocation.append(col[5]) #1 is column index
    tweet.append("DOCID"+str(i)+" "+col[6]) #2 is column index
    i=i+1
print(tweet)