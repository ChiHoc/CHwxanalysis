#!/usr/bin/python
#-*-coding:utf-8-*-

import urllib
import os
import sqlite3
import requests
import json
import xlwt
import math
from Tkinter import *
from datetime import *
from tkMessageBox import showwarning
from pycookiecheat import chrome_cookies

beginDate = datetime.strptime('2017-01-01', '%Y-%m-%d')

wb = xlwt.Workbook()

# cookies = chrome_cookies('http://127.0.0.1')

response = requests.request('GET', 'http://www.baidu.com', headers={}, params={}, cookies=cookies)

respJson = json.loads('{"测试json": true}')
