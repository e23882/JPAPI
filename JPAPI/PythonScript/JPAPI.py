import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re


# Jepun請假API
class JP:
    # 使用者帳號
    username = 'leoyang'
    # 使用者密碼
    password = '12345'
    # 年分
    year = '109'
    # 月份
    month = '09'

    dynamicParameter = ''
    dynamicParameterValue = ''
    globalCookie = ''

    # 建構子
    def __init__(self, username, password, year, month):
        self.username = username
        self.password = password
        self.year = year
        self.month = month

    # 登入
    def Login(self):
        url = "http://secom/work_time/MainN.asp"
        querystring = {"Account": self.username, "password": self.password, "x": "38", "y": "18"}
        headers = {
            'User-Agent': "PostmanRuntime/7.17.1",
            'Accept': "*/*",
            'Cache-Control': "no-cache",
            'Postman-Token': "bcec3a56-5e51-41cf-9f29-60f5829bf74f,70c65962-6021-4061-89f5-affd99ba20ea",
            'Host': "secom",
            'Accept-Encoding': "gzip, deflate",
            'Cookie': "Personal=Password=&Name=; ASPSESSIONIDQABQSSBT=DGEHIADCJKJLIIELCMEKJOHC; ASPSESSIONIDAQDTSRBT=KJAFGCDCJPNFLJPLFIECHBNF",
            'Content-Length': "0",
            'Connection': "keep-alive",
            'cache-control': "no-cache"
        }
        response = requests.request("POST", url, headers=headers, params=querystring)
        globalCookie = response.cookies.get_dict()
        # print(cookies)
        st = str(globalCookie)
        st = st[1:int(st.index(','))].replace("'", "").replace(":", "=").replace(" ", "")
        self.dynamicParameter = st[0:st.index('=')]
        self.dynamicParameterValue = st[st.index('=') + 1:]
        if self.dynamicParameter != '' and self.dynamicParameterValue != '':
            return 1
        else:
            return 0

    # 取得本月差勤資料到暫存資料庫(test.db SQLite)
    def GetThisMonthArriveData(self):
        headers = {
            'Content-Type': "text/xml;charset=UTF-8",
            'User-Agent': "PostmanRuntime/7.17.1",
            'Accept': "*/*",
            'Cache-Control': "no-cache",
            'Postman-Token': "fe30cd2c-cf0b-4344-8f6b-432c00c2be95,67dd0357-6993-4030-91c1-6a01033a640d",
            'Host': "secom",
            'Accept-Encoding': "gzip, deflate",
            'Cookie': "Personal=Password=&Name=; " + self.dynamicParameter + "=" + self.dynamicParameterValue,
            'Connection': "keep-alive",
            'cache-control': "no-cache"
        }

        data = {
            'Effect': '%20',
            'CheckStr': 'OK',
            'Account': self.username,
            'StartYear': self.year,
            'StartMonth': self.month,
            'StartDay': '01',
            'EndYear': self.year,
            'EndMonth': self.month,
            'EndDay': '31',
            'kind': '%20',
            'Ser': 'No',
            'PageCount': '100',
            'image1.x': '41',
            'image1.y': '26'
        }
        response = requests.request("Get", 'http://secom/work_time/Voc/OnTimeQya1.asp', headers=headers, params=data)
        response.encoding = 'big5'
        if response.status_code == 200:
            # print('select ok')
            result = response.text[response.text.index('<html>'):response.text.index('<html>') + response.text.index(
                '</html>')].replace('\r', '').replace('\n', '').replace('\t', '')
            soup = bs(result.replace('s_Clr2', 's_Clr1'), 'html.parser')
            usefulData = soup.find_all('td', 'Rs_Clr1')
            Counter = 0
            data = ''
            self.Delete('delete from WorkTime')
            for item in usefulData:
                if Counter == 0:
                    data = "'" + item.text.strip().replace('時', ':').replace('分', '').replace('  ', '') + "'"
                else:
                    data = data + ",'" + item.text.strip().replace('時', ':').replace('分', '').replace('  ', '') + "'"
                Counter = Counter + 1
                if Counter == 7:
                    Counter = 0
                    self.InsertUpdate(
                        "INSERT INTO WorkTime (RowNo, Name, Department, WeekDay, Arrive, Leave, Other) VALUES (" + data + ");")
                    data = ''
                else:
                    pass
            self.InsertUpdate("update WorkTime set Arrive = '99:99' where Arrive = '無資料';")
            self.InsertUpdate("update WorkTime set Leave = '99:99' where Leave = '無資料';")
            return 1
        else:
            return 0

    # 計算輸出加班時間、明細(from LocalDB)
    def CalculateLeaveDay(self):
        try:
            conn = sqlite3.connect(r"C:\Users\LeoYang\source\repos\JPAPI\JPAPI\bin\Debug\netcoreapp2.1\Test.db")
            cursorObj = conn.cursor()
            cursorObj.execute(
                "SELECT substr(WeekDay, 1, INSTR(WeekDay, '星期')-1) as 'Date', substr(Arrive, 1, 2) as 'ArriveHR', substr(Arrive, 4, 2) as 'ArriveMin', substr(Leave, 1, 2) as 'LeaveHR', substr(Leave, 4, 2) as 'LeaveMin', case when instr(Arrive, '99') = 1 then 'Error' when instr(Leave, '99') = 1 then 'Error' when substr(Arrive, 1, 2) = substr(Leave, 1, 2) then 'Error' else '' end as 'Status' FROM WorkTime")
            rows = cursorObj.fetchall()
            for row in rows:
                arrive = int(row[1]) * 60 + int(row[2])
                leave = int(row[3]) * 60 + int(row[4])
                data = 0
                overwork = 0
                overDetail = ''
                if (leave - arrive - 570) > 60:
                    # 1.計算加班
                    # 1.1計算加班第一小時
                    data = leave - arrive - 630
                    overwork = overwork + 1
                    # 1.2計算加班第一小時後
                    while data >= 45:
                        data = data - 45
                        overwork = overwork + 1
                    # ================================================================================================================
                    # 2.計算加班明細
                    startHr = 9
                    # 2.1 如果有遲到重新設定開始時間
                    if int(row[1]) > 9:
                        startHr = 9
                    else:
                        startHr = int(row[1])
                    commonGoHomeTime = startHr * 60 + int(row[2]) + 90 + 480
                    currentHr = ''
                    currentMin = ''
                    for i in range(overwork):
                        if i == 0:
                            overDetail = overDetail + str(math.floor(commonGoHomeTime / 60)) + ':' + str(
                                commonGoHomeTime % 60) + '~' + str(
                                math.floor((commonGoHomeTime + 60) / 60)) + ':' + str(
                                (commonGoHomeTime + 60) % 60) + '\r\n'
                            currentHr = str(math.floor((commonGoHomeTime + 60) / 60)).zfill(2)
                            currentMin = str((commonGoHomeTime + 60) % 60).zfill(2)
                            commonGoHomeTime = commonGoHomeTime + 60
                        else:
                            overDetail = overDetail + currentHr + ':' + currentMin + '~' + str(
                                math.floor((commonGoHomeTime + 45) / 60)).zfill(2) + ':' + str(
                                (commonGoHomeTime + 45) % 60).zfill(2) + '\r\n'
                            currentHr = str(math.floor((commonGoHomeTime + 45) / 60)).zfill(2)
                            currentMin = str((commonGoHomeTime + 45) % 60).zfill(2)
                            commonGoHomeTime = commonGoHomeTime + 45
                    print(str(row[0]) + str(row[1]) + ':' + str(row[2]) + '~' + str(row[3]) + ':' + str(
                        row[4]) + '\t加班' + str(overwork) + '小時')
                    print(overDetail)
                else:
                    pass
            return 1
        except Exception:
            return 0

    # 請假
    ## year         年(西元年 ex 2020)
    ## month        月
    ## day          日
    ## startHour    開始時(ex 09)
    ## endHour      結束時(ex 100、110、120)
    ## sample       leo.AskLeave('2020','06','02','09','110')
    ##              leo.AskLeave('2020','6','2','09','110')
    def AskLeave(self, year, month, day, startHour, endHour):
        headers = {
            'Content-Type': "text/xml;charset=UTF-8",
            'User-Agent': "PostmanRuntime/7.17.1",
            'Accept': "*/*",
            'Cache-Control': "no-cache",
            'Postman-Token': "fe30cd2c-cf0b-4344-8f6b-432c00c2be95,67dd0357-6993-4030-91c1-6a01033a640d",
            'Host': "secom",
            'Accept-Encoding': "gzip, deflate",
            'Cookie': "Personal=Password=&Name=; " + self.dynamicParameter + "=" + self.dynamicParameterValue,
            'Connection': "keep-alive",
            'cache-control': "no-cache"
        }

        data = {
            'Sv': '800',
            'Sv1': '53',
            'Account': 'leoyang',
            'Voc': '9',
            'StartYear': year,
            'StartMonth': month,
            'StartDay': day,
            'StartHour': startHour,
            'EndYear': year,
            'EndMonth': month,
            'EndDay': day,
            'EndHour': endHour,
            'x': '33',
            'y': '20',
            'CHeckStr': 'OK'
        }
        response = requests.request("POST", 'http://secom/work_time/Voc/VocLoginUD.asp', headers=headers, params=data)
        if response.status_code == 200:
            return 1
        else:
            return 0
        ## year         年

    # 登入工時
    ## year         年(民國年 ex 109)
    ## month        月
    ## day          日
    ## worktime     工時
    ## project      專案代碼
    ## type         工作類型
    ## team         組別
    ## sample       
    ## 登入一天       
    ##      leo.LoginWorkTime('109', '6', 1, 1, '0000099914', '00K6      ', 1)
    ## 登入整月
    ##      for i in range(1,31):
    ##          result = leo.SearchWorkTime('109', '6', str(i))
    ##          #登入工時
    ##          if result!=0:
    ##              leo.LoginWorkTime('109', '6', str(i), str(result), '0000099914', '00K6      ', 1)
    def LoginWorkTime(self, year, month, day, worktime, project, type, team):
        headers = {
            'Content-Type': "text/xml;charset=UTF-8",
            'User-Agent': "PostmanRuntime/7.17.1",
            'Accept': "*/*",
            'Cache-Control': "no-cache",
            'Postman-Token': "fe30cd2c-cf0b-4344-8f6b-432c00c2be95,67dd0357-6993-4030-91c1-6a01033a640d",
            'Host': "secom",
            'Accept-Encoding': "gzip, deflate",
            'Cookie': "Personal=Password=&Name=; " + self.dynamicParameter + "=" + self.dynamicParameterValue,
            'Connection': "keep-alive",
            'cache-control': "no-cache"
        }

        data = {
            'RecDate': '',
            'e': '9',
            'w': '',
            'Year': year,
            'Month': month,
            'Day': day,
            'Team': team,
            'Prj': project,
            'Dem': '',
            'Hs': worktime,
            'Hs_KindNo': type,
            'x': '27',
            'y': '18'
        }
        response = requests.request("POST", 'http://secom/work_time/Wt/WT_LoginUD.asp', headers=headers, params=data)
        if response.status_code == 200:
            return 1
        else:
            return 0

    # 查詢工時
    def SearchWorkTime(self, year, month, day):
        headers = {
            'Content-Type': "text/xml;charset=UTF-8",
            'User-Agent': "PostmanRuntime/7.17.1",
            'Accept': "*/*",
            'Cache-Control': "no-cache",
            'Postman-Token': "fe30cd2c-cf0b-4344-8f6b-432c00c2be95,67dd0357-6993-4030-91c1-6a01033a640d",
            'Host': "secom",
            'Accept-Encoding': "gzip, deflate",
            'Cookie': "Personal=Password=&Name=; " + self.dynamicParameter + "=" + self.dynamicParameterValue,
            'Connection': "keep-alive",
            'cache-control': "no-cache"
        }

        data = {
            'RecDate': '',
            'e': '8',
            'w': '',
            'Year': year,
            'Month': month,
            'Day': day,
            'Team': '1',
            'Prj': '',
            'Dem': '',
            'Hs': '',
            'Hs_KindNo': ''
        }
        response = requests.request("POST", 'http://secom/work_time/Wt/Wt_login.asp', headers=headers, params=data)
        response.encoding = 'big5'
        if response.status_code == 200:
            try:
                result = response.text.index('您本日的工作時數=')
                todayWorktime = response.text[result + 9:result + 14].replace('小', '')
                return int(todayWorktime)
            except Exception:
                return 0
        else:
            return 0

    # (private method)新增/更新資料 
    def __InsertUpdate(self, query):
        conn = sqlite3.connect(r"D:\Backup\LeoYang\Desktop\Project\own\Python\WorkTIme\Test.db")
        cursorObj = conn.cursor()
        cursorObj.execute(query)
        conn.commit()
        conn.close()
        # print(query + ' Successful')

    # (private method)刪除資料 
    def __Delete(self, query):
        conn = sqlite3.connect(r"D:\Backup\LeoYang\Desktop\Project\own\Python\WorkTIme\Test.db")
        cursorObj = conn.cursor()
        cursorObj.execute(query)
        conn.commit()
        conn.close()
        # print(query+' Successful')
