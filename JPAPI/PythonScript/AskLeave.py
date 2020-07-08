import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import JPAPI

# 請假
# 目前只有補修
if len(sys.argv) < 7:
    print('參數數量錯誤，需要參數數量 7，目前 '+str(len(sys.argv)))
else:
    # 初始化物件
    leo = JPAPI.JP(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    # 登入
    loginResult = leo.Login()
    if loginResult == 0:
        print('Login Fail')
    else:
        pass
    if leo.AskLeave(int(sys.argv[3])+1911, sys.argv[4], sys.argv[5],sys.argv[6],sys.argv[7])==1:
        print("Ask Leave Success")
    else :
        print("Ask Leave Fail")
    