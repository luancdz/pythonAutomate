import win32gui
import win32ui
import win32con
from win32api import GetSystemMetrics

class ScreenShot(object):


    def __init__(self):
        
        self.hwnd = win32gui.FindWindow("SAP_FRONTEND_SESSION", None)

    def screenShot(self, soOrigin, width=GetSystemMetrics(0), height=GetSystemMetrics(1)):
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj=win32ui.CreateDCFromHandle(wDC)
        cDC=dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, width, height)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0),(width, height) , dcObj, (0,0), win32con.SRCCOPY)
        dataBitMap.SaveBitmapFile(cDC, str(soOrigin) + '.bmp')
        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())
