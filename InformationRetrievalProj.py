# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 20:23:23 2020

@author: benic and mikeyg
"""

import pandas as pd
import os

import prep as prep
from nltk.tokenize import TweetTokenizer
import nltk
import re
nltk.download('stopwords')
from nltk.corpus import stopwords
from openpyxl import load_workbook
from nltk.stem import PorterStemmer
from spellchecker import SpellChecker

spell = SpellChecker()

tknzr = TweetTokenizer(preserve_case=True, reduce_len=True, strip_handles=True)
ps = PorterStemmer()
cachedStopWords = stopwords.words("english")

dictionary = {}

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
j=2
tweetmap=dict()
geomap=dict()
for col in fileList:
    geoLocation.append(col[5])#1 is column index
    tweet.append(col[6])#2 is column index
    tweetmap[j]=col[6]
    geomap[j]=col[5]
    j=j+1
counter = 2
for i in tweet:
    terms = prep.ngram_tokenizer(text=i)
    print(terms)
    currentTweet = tknzr.tokenize(i.casefold())
    currentTweet = [word for word in currentTweet if word not in cachedStopWords]
    currentTweet = [s.strip('.') for s in currentTweet] 
    currentTweet = [s.replace('.', '') for s in currentTweet]
    currentTweet = [s.strip(':') for s in currentTweet] 
    currentTweet = [s.replace(':', '') for s in currentTweet]
    currentTweet = [s.strip('!') for s in currentTweet] 
    currentTweet = [s.replace('!', '') for s in currentTweet]
    currentTweet = [s.strip('?') for s in currentTweet] 
    currentTweet = [s.replace('?', '') for s in currentTweet]
    currentTweet = [s.strip('\\') for s in currentTweet] 
    currentTweet = [s.replace('\\', '') for s in currentTweet]
    currentTweet = [s.strip('/') for s in currentTweet] 
    currentTweet = [s.replace('/', '') for s in currentTweet]
    currentTweet = [s.strip('(') for s in currentTweet] 
    currentTweet = [s.replace('(', '') for s in currentTweet]
    currentTweet = [s.strip(')') for s in currentTweet] 
    currentTweet = [s.replace(')', '') for s in currentTweet]
    currentTweet = [s.strip('*') for s in currentTweet] 
    currentTweet = [s.replace('*', '') for s in currentTweet]
    currentTweet = [s.strip(',') for s in currentTweet] 
    currentTweet = [s.replace(',', '') for s in currentTweet]
    currentTweet = [s.strip('"') for s in currentTweet] 
    currentTweet = [s.replace('"', '') for s in currentTweet]
    currentTweet = [s.strip('-') for s in currentTweet] 
    currentTweet = [s.replace('-', '') for s in currentTweet]
    currentTweet = [s.strip('+') for s in currentTweet] 
    currentTweet = [s.replace('+', '') for s in currentTweet]
    currentTweet = [s.strip('&') for s in currentTweet] 
    currentTweet = [s.replace('&', '') for s in currentTweet]
    currentTweet = [s.strip('$') for s in currentTweet] 
    currentTweet = [s.replace('$', '') for s in currentTweet]
    currentTweet = [s.strip(';') for s in currentTweet] 
    currentTweet = [s.replace(';', '') for s in currentTweet]
    pattern = '[0-9]'
    currentTweet = [re.sub(pattern, '', i) for i in currentTweet]
    currentTweet = filter (lambda i: len (i) > 2, currentTweet)
    currentTweet = list(filter(None, currentTweet))
    currentTweet = [spell.correction(i) for i in currentTweet]
    
    currentTweet = [ps.stem(i) for i in currentTweet]
    
    
    print(currentTweet)
    

    

