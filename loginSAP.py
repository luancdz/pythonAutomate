import subprocess
import time
import win32com.client
import psutil

class SapGui(object):

    def __init__(self, landscape):

        def getLandscape(landscape):
            switcher = {
                "PAG": "PAG - North American AG Production (SSO)",
                "AG3": "AG3 - North American Ag (SSO) (003)",
                "AG2": "AG2 - North American Ag (SSO) (003)",
            }
            return switcher.get(landscape, False) 

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

        #landscape = getLandscape(landscape)

        self.connection = application.OpenConnection(getLandscape(landscape),True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

    def killSAP(self):
        
        self.connection = None
        self.session = None

        for proc in psutil.process_iter():
            if proc.name() == "saplogon.exe":
                proc.kill()

    def getSession(self):
        return self.session



 
    # get() method of dictionary data type returns
    # value of passed argument if it is present
    # in dictionary otherwise second argument will
    # be assigned as default value of passed argument
        return switcher.get(landscape, False)  


#if __name__ == '__main__':
#    Object  =  SapGui();