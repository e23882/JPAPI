import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import JPAPI

# 請假

# 初始化物件
chris = JPAPI.JP('chrisli', '12345', '109', '6')
# 登入
loginResult = chris.Login()
if loginResult == 0:
    print('Login Fail')
else:
    pass
# if chris.AskLeave('2020','06','02','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','05','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','08','09','110') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','10','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','12','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','15','09','120') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','16','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','17','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','19','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','22','09','110') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','24','09','100') == 1:
#     print('ok')
# if chris.AskLeave('2020','06','29','09','100') == 1:
#     print('ok')
#                   year, month, day, worktime, project, type, team):

for i in range(2,31):
    result = chris.SearchWorkTime('109', '6', str(i))
    print(result)
    #登入工時
    if result!=0:
        chris.LoginWorkTime('109', '6', str(i), str(result), '0000099914', '00K6      ', 29)