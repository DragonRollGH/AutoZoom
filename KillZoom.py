import os
import time

print('将在30s后强制关闭Zoom')
time.sleep(30)
os.system('taskkill /F /T /IM Zoom.exe')
