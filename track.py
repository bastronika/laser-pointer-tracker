#import win32api, win32con
import time
import pyautogui

try:
    while True:
        print pyautogui.position()

except KeyboardInterrupt:
    print('\n')
