import os
import webbrowser
import win32con
import win32api
import time

NetStatusDelay = 10 #打开浏览器x秒后模拟Enter键，视网速情况自行决定，仅测试了Chrome。

"""
Meetings = [
     [0,'https://zoom.com.cn/j/691620952?pwd=WmNLV2xaMC9JdmEwakc4WUNXSmRHdz09',0,0,'https://zoom.com.cn/j/158390142?pwd=ZENBMTMxWTV3c3duT3ZMTm1OSzVxQT09']
    ,['https://zoom.com.cn/j/671315846?pwd=alUzOE8weUV3Y0FIczN2QkFvQTF5dz09',0,0,0,0]
    ,[0,'https://zoom.com.cn/j/709527567?pwd=eGJtamlMVmJJM0Qwc1F1a0h4SlpwZz09',0,'https://zoom.com.cn/j/875144514?pwd=bWpzNEVxNDBEUlBXbFhsV2FOMWtxUT09','https://zoom.com.cn/j/289810142?pwd=VHEwUVplU0cyWWpaRzFYbUw5YmMvQT09']
    ,[0,'https://zoom.com.cn/j/671315846?pwd=alUzOE8weUV3Y0FIczN2QkFvQTF5dz09',0,0,0]
    ,[0,'https://zoom.com.cn/j/910288926?pwd=WnJ0Qng2T1Y1cmJ1L0crbEpJQnlSQT09',0,'https://zoom.com.cn/j/192814831?pwd=bFpyckFKd3htOUlFdWdId3k3c3o5QT09',0]
]

ifRecord = [
     [0,0,0,0,0]
    ,[0,0,0,0,0]
    ,[0,0,0,0,0]
    ,[0,0,0,0,0]
    ,[0,0,0,0,0]
]
"""
### 学期结束了，所以注释掉了上面的课程表，下面是小学期的课程，每天都是同一个链接
Meetings = 'https://zoom.com.cn/j/65372678730?pwd=NDVYR3RjdHBwQmxpTmxqZWk5ZzY3Zz09'

def Logging(txt):
    with open('log.txt','a') as log:
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        log.write(t+'\t'+txt+'\n')

def JoinMeeting(Url,DelayRate=1):
    webbrowser.open(Url)
    time.sleep(NetStatusDelay*DelayRate)

    win32api.keybd_event(9, 0, 0, 0)  # 键盘按下Tab
    time.sleep(0.5)
    win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)  # 键盘松开
    time.sleep(0.5)
    win32api.keybd_event(13, 0, 0, 0)  # 键盘按下Enter
    time.sleep(0.5)
    win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)  # 键盘松开

######判断星期及第几节课######
Date = time.localtime()
Week = Date[6]
Num = 5
if Date[3] == 7 and 50 <= Date[4] <= 59:
    Num = 0

elif Date[3] == 9 and 50 <= Date[4] <= 59:
    Num = 1

elif Date[3] == 13 and 20 <= Date[4] <= 29:
    Num = 2

elif Date[3] == 15 and 25 <= Date[4] <= 29:
    Num = 3

elif Date[3] == 18 and 50 <= Date[4] <= 59:
    Num = 4
######判断星期及第几节课######

# if Num == 5 or Week in (5,6):             #端午节调休临时注释掉了
#     Logging('不在上课时间')
#     quit()

#CMeeting = Meetings[Week][Num]             #端午节调休临时注释掉了
CMeeting = Meetings

if not CMeeting:
    Logging('现在没课')
    quit()

while True:

    DelayRate = 1
    JoinMeeting(CMeeting, DelayRate)
    DelayRate += 1
    time.sleep(30)

    if  "Zoom.exe" in os.popen('tasklist /FI "IMAGENAME eq Zoom.exe"').read():
        Logging('成功打开Zoom')
        win32api.keybd_event(120, 0, 0, 0)  # 键盘按下F9
        time.sleep(0.5)
        win32api.keybd_event(120, 0, win32con.KEYEVENTF_KEYUP, 0)  # 键盘松开
        break

    if DelayRate >= 10:
        Logging('Zoom打开失败')
        break


