# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 20:23:23 2020

@author: benic and mikeyg
"""
"TO USE NLTK GO TO PYTHON CONSOLE AND WRITE THE FOLLOWING COMMANDS" \
"import nltk" \
"ntlk.download()"
import pandas as pd
import os

from openpyxl import load_workbook
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize


def iter_rows(ws):
    result = []
    for row in ws.iter_rows():
        rowlist = []
        for cell in row:
            rowlist.append(cell.value)
        result.append(rowlist)
    return result


workbook = load_workbook(filename='data.xlsx', read_only=True)
worksheet = workbook['sheet']
fileList = (list(iter_rows(worksheet)))
geoLocation = []
tweet = []
i = 2
for col in fileList:
    geoLocation.append(col[5])  # 1 is column index
    tweet.append("DOCID" + str(i) + " " + col[6])  # 2 is column index
    i = i + 1
stop_words = set(stopwords.words('english'))
word_tokens = word_tokenize(str(tweet))
filtered_sentence = []
current_doc = ''
diction = dict()
for w in word_tokens:
    if w not in stop_words:
        if "DOCID" in w:
            current_doc = w
        else:
            #diction.update(w=diction.get(w).append(current_doc))
    #TODO finish creating dictionary
