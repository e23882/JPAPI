import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import JPAPI

# 計算加班時數
if len(sys.argv) < 4:
    print('參數數量錯誤')
else:
    # 初始化物件
    leo = JPAPI.JP(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    # 登入
    loginResult = leo.Login()
    if loginResult == 0:
        print('Login Fail')
    else:
        pass
    # 取得當月差勤資料
    getDataResult = leo.GetThisMonthArriveData()
    if getDataResult == 0:
        print('Failed to get the attendance data of the month')
        sys.exit()
    else:
        pass

    # 計算差勤輸出請假資料
    calculateResult = leo.CalculateLeaveDay()
    if calculateResult == 0:
        print('Calculate Fail')
    else:
        pass
