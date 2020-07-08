---
title: JPAPI
catalog: true
date: 2020-07-02 20:25:17
subtitle:
header-img:
tags:
- Python
- API
- Company
---
# JPAPI
Company's internal system for taking leave and registering hours python API
### Login － 登入
登入系統。所有Request中，cookie中都會需要一個DynamicParameter，名稱跟值都是動態的，DynamicParameter有時候會跟舊的一樣，有時候會跟新的一樣，建議每次都重新登入取得最新的DynamicParameter避免送Request失敗
### GetThisMonthArriveData － 取得本月差勤資料到本地端的資料庫
GetThisMonthArriveData會去爬目前最新的差勤資料回來，存在SQLite中，提供給計算加班時間明細(CalculateLeaveDay)使用
### CalculateLeaveDay － 計算、輸出加班時間明細
CalculateLeaveDay會從本地資料庫中取得當月上班時間，計算當月加班時數及加班明細
### AskLeave － 請假
就是請假，目前只能請補修因為我懶惰，請假類型應該是送出Request Data中的'Voc'欄位
### SearchWorkTime － 查詢工時
team : 1  業務部
team : 29 研發部

### LoginWorkTime － 登入工時
就是登工時，可以搭配查詢工時使用一次查詢當月1:31號的工時。然後再一次登入Hen方便




# QuickStart
## WebAPI
### Python - Crawler
- Python 3.7
- requests (python lib)
- BeautifulSoup (python lib)
- sqlite3 (python lib)
### C# - WebAPI
- .NetCore2.1
***
