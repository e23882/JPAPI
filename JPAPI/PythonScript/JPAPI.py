import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re


# Jepun請假API

# (static method)新增/更新資料
def InsertUpdate(query):
    conn = sqlite3.connect(r".\Test.db")
    cursorObj = conn.cursor()
    cursorObj.execute(query)
    conn.commit()
    conn.close()
    # print(query + ' Successful')


# (static method)刪除資料
def Delete(query):
    conn = sqlite3.connect(r".\Test.db")
    cursorObj = conn.cursor()
    cursorObj.execute(query)
    conn.commit()
    conn.close()


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
    # username 使用者名稱
    # password 使用者密碼
    # yaer     西元年
    # month    月分(要存取資料的月份)
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
            Delete('delete from WorkTime')

            for item in usefulData:
                if Counter == 0:
                    data = "'" + item.text.strip().replace('時', ':').replace('分', '').replace('  ', '') + "'"
                else:
                    data = data + ",'" + item.text.strip().replace('時', ':').replace('分', '').replace('  ', '') + "'"
                Counter = Counter + 1
                if Counter == 7:
                    Counter = 0
                    InsertUpdate(
                        "INSERT INTO WorkTime (RowNo, Name, Department, WeekDay, Arrive, Leave, Other) VALUES (" + data + ");")
                    data = ''
                else:
                    pass
            InsertUpdate("update WorkTime set Arrive = '99:99' where Arrive = '無資料';")
            InsertUpdate("update WorkTime set Leave = '99:99' where Leave = '無資料';")
            return 1
        else:
            return 0

    # 計算請假後加班時間、明細(from LocalDB)
    def OvertimeAfterLeaveOfAbsence(self):
        try:
            conn = sqlite3.connect(r".\Test.db")
            cursorObj = conn.cursor()
            cursorObj.execute("SELECT substr(WeekDay, 1, INSTR(WeekDay, '星期')-1) as 'Date', substr(Arrive, 1, 2) as 'ArriveHR', substr(Arrive, 4, 2) as 'ArriveMin', substr(Leave, 1, 2) as 'LeaveHR', substr(Leave, 4, 2) as 'LeaveMin', case when instr(Arrive, '99') = 1 then 'Error' when instr(Leave, '99') = 1 then 'Error' when substr(Arrive, 1, 2) = substr(Leave, 1, 2) then 'Error' else '' end as 'Status' FROM WorkTime")
            rows = cursorObj.fetchall()
            for row in rows:
                LocalStartHr = '09'
                LocalStartMin = '00'
                LocalEndHr = '18'
                LocalEndMin = '30'
                arrive = 0
                # 上班打卡 時
                LocalStartHr = str(row[1])
                # 上班打卡 分
                LocalStartMin = str(row[2])
                # 下單打卡 時
                LocalEndHr = str(row[3])
                # 下單打卡 分
                LocalEndMin = str(row[4])
                # 上班時間換算成分鐘(用在計算是否滿足加班)
                arrive = 0
                # 應該請假時數
                LeaveAbsence = 0

                # 判斷是否遲到，遲到的話 計算遲到時數，設定上班時間為九點
                if int(LocalStartHr) == 10:
                    LeaveAbsence = 1
                    LocalStartHr = '09'
                elif int(LocalStartHr) == 11:
                    LeaveAbsence = 2
                    LocalStartHr = '09'
                elif int(LocalStartHr) == 12:
                    LeaveAbsence = 3
                    LocalStartHr = '09'
                    LocalStartMin = '00'
                elif int(LocalStartHr) == 13:
                    LeaveAbsence = 3
                    LocalStartHr = '09'
                    if int(LocalStartMin) < 30:
                        LocalStartMin = '00'
                    else:
                        pass
                elif int(LocalStartHr) == 14:
                    LeaveAbsence = 4
                    LocalStartHr = '09'
                    if int(LocalStartMin) < 30:
                        LocalStartMin = '00'
                    else:
                        pass
                else:
                    pass
                arrive = int(LocalStartHr) * 60 + int(LocalStartMin)
                leave = int(LocalEndHr) * 60 + int(LocalEndMin)
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
                    if int(LocalStartHr) > 9:
                        startHr = 9
                    else:
                        startHr = int(LocalStartHr)
                    commonGoHomeTime = startHr * 60 + int(LocalStartMin) + 90 + 480
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
                    if LeaveAbsence > 0:
                        print(str(row[0]) + str(row[1]) + ':' + str(row[2]) + '~' + str(row[3]) + ':' + str(
                            row[4]) + '\t請假' + str(LeaveAbsence) + '小時' + '\t加班' + str(overwork) + '小時')
                    else:
                        print(str(row[0]) + str(row[1]) + ':' + str(row[2]) + '~' + str(row[3]) + ':' + str(row[4]) + '\t加班' + str(
                            overwork) + '小時')
                    print(overDetail)
                else:
                    pass
            return 1
        except Exception:
            return 0

    # 計算請假前加班時間、明細(from LocalDB)
    def CalculateLeaveDay(self):
        try:
            conn = sqlite3.connect(r".\Test.db")
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

    def ShowWorkData(self):
        try:
            conn = sqlite3.connect(r".\Test.db")
            cursorObj = conn.cursor()
            cursorObj.execute(
                "SELECT substr(WeekDay, 1, INSTR(WeekDay, '星期')-1) as 'Date', substr(Arrive, 1, 2) as 'ArriveHR', substr(Arrive, 4, 2) as 'ArriveMin', substr(Leave, 1, 2) as 'LeaveHR', substr(Leave, 4, 2) as 'LeaveMin', case when instr(Arrive, '99') = 1 then 'Error' when instr(Leave, '99') = 1 then 'Error' when substr(Arrive, 1, 2) = substr(Leave, 1, 2) then 'Error' else '' end as 'Status' FROM WorkTime")
            rows = cursorObj.fetchall()
            for row in rows:
                print(str(row[0]) + str(row[1]) + ':' + str(row[2]) + '~' + str(row[3]) + ':' + str(row[4]) + '\t' + str(row[5]))
            return 1
        except Exception:
            return 0

    # 請假
    # year         年(西元年 ex 2020)
    # month        月
    # day          日
    # startHour    開始時(ex 09)
    # endHour      結束時(ex 100、110、120)
    # sample       leo.AskLeave('2020','06','02','09','110')
    #              leo.AskLeave('2020','6','2','09','110')
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
    # year         年(民國年 ex 109)
    # month        月
    # day          日
    # worktime     工時
    # project      專案代碼
    # type         工作類型
    # team         組別
    # sample
    # 登入一天
    #      leo.LoginWorkTime('109', '6', 1, 1, '0000099914', '00K6      ', 1)
    # 登入整月
    #      for i in range(1,31):
    #          result = leo.SearchWorkTime('109', '6', str(i))
    #          #登入工時
    #          if result!=0:
    #              leo.LoginWorkTime('109', '6', str(i), str(result), '0000099914', '00K6      ', 1)
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
                todayWorktime = response.text[result + 9:result + 14].replace('小時', '').replace('時', '').replace('小', '').replace('.0', '')
                return int(todayWorktime)
            except Exception:
                return 0
        else:
            return 0

# sample code
# if __name__ == '__main__':

    # leo = JP('leoyang', '12345', '109', '07')
    # # 登入
    # loginResult = leo.Login()
    # if loginResult == 0:
    #     print('登入失敗')
    #     sys.exit()
    # else:
    #     print('登入成功')

    # 取得當月差勤資料
    # getDataResult = leo.GetThisMonthArriveData()
    # if getDataResult == 0:
    #     print('取得當月差勤資料失敗')
    #     sys.exit()
    # else:
    #     print('取得當月差勤資料成功')

    #印出工作資料
    leo.ShowWorkData()
    #印出加班請假資料
    leo.OvertimeAfterLeaveOfAbsence()

    #請假(請假要自己算)
    # leo.AskLeave('2020', '07', '03', '09', '120')
    # leo.AskLeave('2020', '07', '06', '09', '120')
    # leo.AskLeave('2020', '07', '07', '09', '100')
    # leo.AskLeave('2020', '07', '08', '183', '193')
    # leo.AskLeave('2020', '07', '10', '09', '100')
    # leo.AskLeave('2020', '07', '13', '09', '110')   #應該沒早退
    # leo.AskLeave('2020', '07', '14', '09', '100')
    # leo.AskLeave('2020', '07', '15', '09', '110')
    # leo.AskLeave('2020', '07', '17', '09', '100')
    # leo.AskLeave('2020', '07', '20', '09', '100')
    # leo.AskLeave('2020', '07', '27', '09', '120')
    # leo.AskLeave('2020', '07', '28', '09', '100')
    # leo.AskLeave('2020', '07', '29', '09', '100')
    # leo.AskLeave('2020', '07', '30', '09', '100')
    # leo.AskLeave('2020', '07', '31', '09', '110')

    # 登工時
    # for i in range(1, 31):
    #     result = leo.SearchWorkTime('109', '7', str(i))
    #     # 登入工時
    #     if result != 0:
    #         print('登入 7/' + str(i) + ' 工時' + str(result))
    #         leo.LoginWorkTime('109', '7', str(i), str(result), '0000099914', '00K6      ', 1)
