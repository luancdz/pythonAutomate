#import win32com.client
import sys
import openpyxl
import time

from win32com.client.makepy import main
import loginSAP
from   runScripstClass import RunScripts
import constant
from tkinter import *
from tkinter import messagebox

def SAP_OP():

    #call class open SAP GUI
    sapGui  =  loginSAP.SapGui()
    session = sapGui.session 
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/use/dsm"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btnBUT_SYNC_EXPORT").press()
#    session.findById("wnd[0]/usr/tabsG_TAB_0200_OBJ_SEL/tabpMAIN_OBJ/ssubSUB_SELECTION:/USE/SAPLPDS3:(0201)/cntlMAIN_SELECTION_CONTAINER/shellcont/shell/shellcont[1]/shell").hierarchyHeaderWidth = 1250
    session.findById("wnd[0]/usr/tabsG_TAB_0200_OBJ_SEL/tabpTEMPLATES").select()
    session.findById("wnd[0]/usr/tabsG_TAB_0200_OBJ_SEL/tabpTEMPLATES/ssubSUB_SELECTION:/USE/SAPLPDS3:0202/cntlTEMP_SELECTION_CONTAINER/shellcont/shell/shellcont[1]/shell").selectedNode = "          2"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[8]").select()
    session.findById("wnd[1]/usr/ctxtGS_0246-FILENAME").text = "C:\\2021\Automation\SO.csv"
    session.findById("wnd[1]/usr/ctxtGS_0246-FILENAME").caretPosition = 25
    session.findById("wnd[1]/usr/btnUPT").press()
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
if __name__ == '__main__':
    SAP_OP()