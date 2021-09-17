import subprocess
import time
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
        SapGuiAuto = None

        #login AG2
        self.connection = application.OpenConnection("AG2 - North American Ag",True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "VMTESTID"
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "yellow2021"
        self.session.findById("wnd[0]").sendVKey(0)

    def killSAP(self):
        
        self.connection = None
        self.session = None

        for proc in psutil.process_iter():
            if proc.name() == "saplogon.exe":
                proc.kill()

    def getSession(self):
        return self.session


#if __name__ == '__main__':
#    Object  =  SapGui();