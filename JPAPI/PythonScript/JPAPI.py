import requests
import sys
from bs4 import BeautifulSoup as bs
import sqlite3
import os.path
import math
import re
import math

# Jepun請假API

# (static method)新增/更新資料
def InsertUpdate(query):
    conn = sqlite3.connect(r".\Test.db")
    cursorObj = conn.cursor()
    cursorObj.execute(query)
    conn.commit()
    conn.close()

# (static method)刪除資料
def Delete(query):
    conn = sqlite3.connect(r".\Test.db")
    cursorObj = conn.cursor()
    cursorObj.execute(query)
    conn.commit()
    conn.close()

class JP:
    # Region Declarations
    # 使用者帳號
    username = 'leoyang'
    # 使用者密碼
    password = '12345'
    # 年分
    year = '109'
    # 月份
    month = '09'
    
    #JWT
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

    # 計算加班、請假明細 & 當月加班請假差
    def OvertimeAfterLeaveOfAbsence(self):
        conn = sqlite3.connect(r".\Test.db")
        cursorObj = conn.cursor()
        cursorObj.execute("SELECT substr(WeekDay, 1, INSTR(WeekDay, '星期')-1) as 'Date', substr(Arrive, 1, 2) as 'ArriveHR', substr(Arrive, 4, 2) as 'ArriveMin', substr(Leave, 1, 2) as 'LeaveHR', substr(Leave, 4, 2) as 'LeaveMin', case when instr(Arrive, '99') = 1 then 'Error' when instr(Leave, '99') = 1 then 'Error' when substr(Arrive, 1, 2) = substr(Leave, 1, 2) then 'Error' else '' end as 'Status' FROM WorkTime")
        rows = cursorObj.fetchall()
        count = 0
        for row in rows:
            # 日期
            date = row[0]
            
            # 上班時間(時)
            arriveHr = int(row[1])
            
            # 上班時間(分)
            arriveMin= int(row[2])
            
            # 下班時間(時)
            leavHr = int(row[3])
            
            # 下班時間(分)
            leavMin = int(row[4])
            
            # 取得上班遲到時數(時)
            LateResult = self.__GetLateHour(arriveHr, arriveMin, leavHr, leavMin)
            # 取得加班時數(時)
            OverworkResult = self.__GetWorkOver(arriveHr, arriveMin, leavHr, leavMin, LateResult)
            # 取得加班明細
            detail = self.__GetWorkoverDetail(arriveHr, arriveMin, leavHr, leavMin, LateResult, OverworkResult)
            # 取得早退時數(時)
            earlyLeave = self.__GetEarlyLeave(arriveHr, arriveMin, leavHr, leavMin, LateResult)
            # 加班請假差(時)
            count = count - LateResult + OverworkResult -earlyLeave
            
            print(f"{date} {str(arriveHr).zfill(2)}:{str(arriveMin).zfill(2)} {str(leavHr).zfill(2)}:{str(leavMin).zfill(2)}\t遲到 {LateResult} 小時\t加班 {OverworkResult} 小時 \t早退{earlyLeave} 小時")
            if detail != '':
                print(detail)
  
        if count < 0 :
            print(f"本月額外請假時數 {-count}小時")
        else :
            print(f"本月額外補修時數 {count}小時")
            
    # 取得遲到時數(Private Function)
    def __GetLateHour(self, arriveHr, arriveMin, leaveHr, leaveMin):
        # 10:0X遲到
        if arriveHr == 10 and arriveMin > 0:
            return 1
        # 11:0x遲到
        elif arriveHr == 11 and arriveMin > 0:
            return 2
        elif (arriveHr == 12 and arriveMin > 0) or (arriveHr==13 and arriveMin<=30):
            return 3
        elif arriveHr == 13 and arriveMin > 30:
            return 4
        elif arriveHr == 14 and arriveMin > 30:
            return 5
        elif arriveHr == 15 and arriveMin > 30:
            return 6
        elif arriveHr == 16 and arriveMin > 30:
            return 7
        elif arriveHr == 17 and arriveMin > 30:
            return 8
        else:
            return 0
    
    # 取得加班時數(Private Function)
    def __GetWorkOver(self, arriveHr, arriveMin, leaveHr, leaveMin, lateHour):
        # 打卡記錄轉換成分鐘數方便計算
        totalWorktime = (leaveHr*60+leaveMin)-(arriveHr*60+arriveMin)
        # 扣掉午休
        totalWorktime = totalWorktime - 90
        # 扣掉八小時上班時間
        totalWorktime = totalWorktime - 480
        # 扣掉遲到修正時數
        totalWorktime = totalWorktime + lateHour*60
        
        WorkoverHour = 0
        while totalWorktime >60:
            totalWorktime -= 60
            WorkoverHour = WorkoverHour+1
            
        if totalWorktime>45:
            totalWorktime -= 45
            WorkoverHour = WorkoverHour+1
            
        return WorkoverHour
   
    # 取得加班明細(Private Function)
    def __GetWorkoverDetail(self, arriveHr, arriveMin, leaveHr, leaveMin, lateHour, overworkHour):
        result = ''
        arrHr = arriveHr
        arrMin = arriveMin
        levHr = leaveHr
        levMin = leaveMin
        
        baseLeaveHr = 18
        baseLeaveMin = 30
        # 如果有遲到，計算理論下班時間
        if lateHour >0:
            baseLeaveMin = baseLeaveMin+arriveMin
            if baseLeaveMin > 60:
                baseLeaveMin = baseLeaveMin - 60
                baseLeaveHr = baseLeaveHr +1
        
        # 如果沒有加班，返回空字串
        if overworkHour == 0:
            return result
            
        if overworkHour == 1:
            return f"{str(baseLeaveHr).zfill(2)}:{str(baseLeaveMin).zfill(2)} ~ {str(baseLeaveHr+1).zfill(2)}:{str(baseLeaveMin).zfill(2)}\r\n"
        # 如果有加班，計算加班明細，
        else :
            for i in range(0,overworkHour):
                
                if overworkHour-i == 1:
                    result += f"{str(baseLeaveHr).zfill(2)}:{str(baseLeaveMin).zfill(2)} ~ {str(baseLeaveHr+1).zfill(2)}:{str(baseLeaveMin-15).zfill(2)}\r\n"
                else:
                    result += f"{str(baseLeaveHr).zfill(2)}:{str(baseLeaveMin).zfill(2)} ~ {str(baseLeaveHr+1).zfill(2)}:{str(baseLeaveMin).zfill(2)}\r\n"
                    baseLeaveHr = baseLeaveHr+1
                    
        return result
    
    # 取得早退時數
    def __GetEarlyLeave(self, arriveHr, arriveMin, leaveHr, leaveMin, lateHour):
        arvHr = arriveHr
        arvMin = arriveMin
        levHr= leaveHr
        levMin = leaveMin
        # 打卡時間轉換分鐘，計算工時
        acturalWork = (levHr*60+levMin) - (arvHr*60+arvMin) - (90+480)+ (lateHour*60)
        if acturalWork<0:
            return math.ceil((-acturalWork)/60)
        else:
            return 0
        
    
    # 請假
    # year         年(西元年 ex 2020)
    # month        月
    # day          日
    # startHour    開始時(ex 09)
    # endHour      結束時(ex 100、110、120)
    #   Sample       leo.AskLeave('2020','06','02','09','110')
    #                leo.AskLeave('2020','6','2','09','110')
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
            print(f"{month}/{day} 請假成功")
            return 1
        else:
            print(f"{month}/{day} 請假失敗")
            return 0
        
    # 登入工時
    # year         年(民國年 ex 109)
    # month        月
    # day          日
    # worktime     工時
    # project      專案代碼
    # type         工作類型
    # team         組別(1.業務部 2.開發  3.客服  4.顧問       5.管理       6.新進人員  7.離職    8.外包人力  9.國泰駐外  10.測試部
    #                   11.技術  12.任務 13.顧服 14.技術一部  15.技術二部  16.技術三部 17.專案組 18.技術股務 19.技術投信 20.技術證券
    #                   21.專案三組 22.專案一組 23.專案二組 24.技術銀行一 25.技術銀行二 26.專案測試部 27.專案部 28.總經理室 29.研發部 30.管理顧問
    #                   31.技術支援 32.銀行證券 33.技術三組 34.技術五組 35.海外事業部
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
            print(f"登入 {month}/{day} 工時成功")
            return 1
        else:
            print(f"登入 {month}/{day} 工時失敗")
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
    
    # 檢查是否有異常打卡資料(沒打上班卡、沒打下班卡)
    def CheckError(self):
        result = 0
        conn = sqlite3.connect(r".\Test.db")
        cursorObj = conn.cursor()
        cursorObj.execute("SELECT Arrive FROM WorkTime where Arrive = 99")
        rows = cursorObj.fetchall()
        if len(rows) > 0:
            return 0
        else:
            return 1



