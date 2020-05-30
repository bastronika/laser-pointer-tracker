import win32api, win32con
import time

def click():
    
    for i in range(720):
        x = int(i*1.5)
        y = i
        win32api.SetCursorPos((x,y))
        time.sleep(0.01)
    
click()