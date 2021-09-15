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
    session.findById("wnd[0]").resizeWorkingPane(141,26,FALSE)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nsm30"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtVIEWNAME").text = "ZPAT_T_FT"
    session.findById("wnd[0]/usr/ctxtVIEWNAME").caretPosition = 9
    session.findById("wnd[0]/usr/btnSHOW_PUSH").press
if __name__ == '__main__':
    SAP_OP()