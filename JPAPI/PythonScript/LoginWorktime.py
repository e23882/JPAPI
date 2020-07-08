import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import JPAPI

#登工時
if len(sys.argv) < 8:
    print('參數數量錯誤'+str(len(sys.argv)))
else:
    # 初始化物件
    leo = JPAPI.JP(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    # 登入
    loginResult = leo.Login()
    if loginResult == 0:
        print('Login Fail')
    else:
        pass
    if leo.LoginWorkTime( sys.argv[3],  sys.argv[4], str(sys.argv[5]), str(sys.argv[6]), sys.argv[7], sys.argv[8], sys.argv[9]) == 1:
        print("Login Worktime Success")
    else :
        print("Login Worktime Fail")
    