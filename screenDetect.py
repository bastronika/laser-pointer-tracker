from win32api import GetSystemMetrics

print("Width =", GetSystemMetrics(0))
print("Height =", GetSystemMetrics(1))