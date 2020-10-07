import JPAPI


#chris = JP('chrisli', '12345', '109', '09')
leo = JPAPI.JP('leoyang', '12345', '109', '09')


# 登入
loginResult = leo.Login()
#chris.Login()
if loginResult == 0:
    print('登入失敗')
    sys.exit()
else:
    print('登入成功')

# 取得當月差勤資料
getDataResult = leo.GetThisMonthArriveData()
if getDataResult == 0:
    print('取得當月差勤資料失敗')
    sys.exit()
else:
    print('取得當月差勤資料成功')
    



#印出加班請假資料
leo.OvertimeAfterLeaveOfAbsence()



# 其實可以自動化，等請假資料相關數據計算更純熟再做

# #請假(請假要自己算)
#leo.AskLeave('2020', '09', '01', '09', '100')
#leo.AskLeave('2020', '09', '02', '09', '100')
#leo.AskLeave('2020', '09', '03', '09', '100')
#leo.AskLeave('2020', '09', '07', '09', '110')
#leo.AskLeave('2020', '09', '08', '09', '100')
#leo.AskLeave('2020', '09', '14', '09', '120')
#leo.AskLeave('2020', '09', '15', '09', '100')
#leo.AskLeave('2020', '09', '17', '145', '195')
#leo.AskLeave('2020', '09', '22', '09', '100')
#leo.AskLeave('2020', '09', '23', '09', '100')
#leo.AskLeave('2020', '09', '23', '185', '195')
#leo.AskLeave('2020', '09', '24', '09', '100')
#leo.AskLeave('2020', '09', '26', '155', '195')
#leo.AskLeave('2020', '09', '28', '09', '100')
#leo.AskLeave('2020', '09', '29', '09', '100')
#leo.AskLeave('2020', '09', '30', '09', '100')


# ErrorDataExists = leo.CheckError()
# if ErrorDataExists == 1:
    # # 登工時
    # for i in range(1, 32):
        # result = leo.SearchWorkTime('109', '9', str(i))
        # # 登入工時
        # if result != 0:
            # leo.LoginWorkTime('109', '9', str(i), str(result), '0000099914', '00K6      ', 14)
            # #chris.LoginWorkTime('109', '8', str(i), str(result), '0000099914', '00K6      ', 14)
# else:
    # print('有異常打卡資料')
    