import win32api, win32con, win32gui
#import pyautogui

def tracking():
    flags, hcursor, (x,y) = win32gui.GetCursorInfo()
    print(x,y)
    
try:
    while True:
        #print pyautogui.position()
        tracking()

except KeyboardInterrupt:
    print('\n')
