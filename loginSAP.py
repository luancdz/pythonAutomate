import subprocess
import time
from types import new_class
import win32com.client

class SapGui(object):



    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(3)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine

        #login AG3
        self.connection = application.OpenConnection("AG3 - North American Ag (SSO) (003)",True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()


#if __name__ == '__main__':
#    Object  =  SapGui();