import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import JPAPI

#取得指定日期工時
if len(sys.argv) < 5:
    print('參數數量錯誤')
else:
    leo = JPAPI.JP(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    loginResult = leo.Login()
    if loginResult == 0:
        print('Login Fail')
    else:
        #print(leo.SearchWorkTime('108', '06', '06'))
        print(leo.SearchWorkTime(sys.argv[3], sys.argv[4], sys.argv[5]))