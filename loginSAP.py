import subprocess
import time
from types import new_class
import win32com.client
import psutil

class SapGui(object):

    def __init__(self):

        for proc in psutil.process_iter():
            if proc.name() == "saplogon.exe":
                proc.kill()


        try:
            #verifica se o SAP GUI ja esta aberto
            sapgui = win32com.client.GetObject("SAPGUI")
            #sapgui.
        except:
            #caso não, abre via .exe
            self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            subprocess.Popen(self.path)
            time.sleep(3)
            

        finally:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")

        #self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine

        #login AG3
        self.connection = application.OpenConnection("AG3 - North American Ag (SSO) (003)",True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()


#if __name__ == '__main__':
#    Object  =  SapGui();