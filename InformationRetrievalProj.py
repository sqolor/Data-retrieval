# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 20:23:23 2020

@author: benic
"""
import pandas as pd
import os

from openpyxl import load_workbook

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

for col in fileList:
    geoLocation.append(col[5])#1 is column index
    tweet.append(col[6])#2 is column index
    

    

