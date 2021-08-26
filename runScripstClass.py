import subprocess
import time
import win32com.client
import sys

class RunScripts(object):


    def __init__(self, conn):
        
        self.SapGuiSession = conn.session

    def getOVfromProd(self, ovOrig):

        try:

            self.SapGuiSession.findById("wnd[0]").maximize()
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtGD-TAB").text = "/USE/PDS3_OBJMAP"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = "PAG"
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "410"
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,3]").text = "SALES_DOC"
            #print(ovOrig.value)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = "*" + str(ovOrig) + "*"
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").caretPosition = 1
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(8)
            self.SapGuiSession.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").currentCellColumn = ""
            self.SapGuiSession.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectedRows = "0"
            self.SapGuiSession.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarButton("DETAIL")
            self.SapGuiSession.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").setFocus()
            self.SapGuiSession.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").caretPosition = 0
            
            return str(self.SapGuiSession.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").text)
        
        except:
            
            raise print(sys.exc_info()[0])




