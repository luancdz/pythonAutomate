import win32com.client
import sys
import openpyxl
import time
import loginSAP

def SAP_OP():

    excelPath = r'#'

    #call class open SAP GUI
    sapGui  =  loginSAP.SapGui()


    # SapGuiAuto = win32com.client.GetObject("SAPGUI")
    # if not type(SapGuiAuto) == win32com.client.CDispatch:
    #     return

    # application = SapGuiAuto.GetScriptingEngine
    # if not type(application) == win32com.client.CDispatch:
    #     SapGuiAuto = None
    #     return

    # connection = application.Children(0)
    # if not type(connection) == win32com.client.CDispatch:
    #     application = None
    #     SapGuiAuto = None
    #     return

    # session = connection.Children(0)
    # if not type(session) == win32com.client.CDispatch:
    #     connection = None
    #     application = None
    #     SapGuiAuto = None
    #     return

    pathExel = "Automate_NF_Transicao_Teste.xlsx"
    wb_obj = openpyxl.load_workbook(pathExel) 
    sheet_obj = wb_obj.active 


    ovOrig = sheet_obj.cell(row = 2, column = 1)  

    #script para pegar o numero da OV do
    sapGui.session.findById("wnd[0]").maximize()
    sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/ctxtGD-TAB").text = "/USE/PDS3_OBJMAP"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = "PAG"
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "410"
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,3]").text = "SALES_DOC"
    print(ovOrig.value)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = "*" + str(ovOrig.value) + "*"
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").caretPosition = 1
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]").sendVKey(8)
    sapGui.session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").currentCellColumn = ""
    sapGui.session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectedRows = "0"
    sapGui.session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarButton("DETAIL")
    sapGui.session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").setFocus()
    sapGui.session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").caretPosition = 0
    ovNew = sapGui.session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").text

    tpOV = sheet_obj.cell(row = 2, column = 10) 
    ovRef = sheet_obj.cell(row = 2, column = 12)

    #criar OV com referencia
    try:
        sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = tpOV.value
        sapGui.session.findById("wnd[0]").sendVKey(8)
        sapGui.session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF").select()
        sapGui.session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").text = ovNew
        sapGui.session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").caretPosition = 10
        sapGui.session.findById("wnd[1]").sendVKey(5)
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = " "
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").setFocus
        sapGui.session.findById("wnd[0]/tbar[0]/btn[11]").press
        sapGui.session.findById("wnd[0]").sendVKey(11)
        sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
        sapGui.session.findById("wnd[0]").sendVKey(0)
        ovRef.value = sapGui.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
    except:
        print(sys.exc_info()[0])
    
    #finally:
        #TBD
 
    #Quantidade de itens da OV
    shipPoint = sheet_obj.cell(row = 2, column = 7)
    wb_obj.save(pathExel)

    sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/ctxtGD-TAB").text = "vbap"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = str(ovRef.value)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").setFocus()
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").caretPosition = 10
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").setFocus()
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").caretPosition = 5
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 3
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 6
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 9
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 12
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 15
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 18
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 21
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 24
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 27
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 30
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 33
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 36
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 39
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 42
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 45
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").text = str(shipPoint.value)
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").setFocus
    sapGui.session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").caretPosition = 4
    sapGui.session.findById("wnd[0]").sendVKey(8)
    sapGui.session.findById("wnd[0]/usr/txtGD-NUMBER").setFocus
    sapGui.session.findById("wnd[0]/usr/txtGD-NUMBER").caretPosition = 0
    itmnun = sapGui.session.findById("wnd[0]/usr/txtGD-NUMBER").text
    sapGui.session.findById("wnd[0]/tbar[0]/btn[3]").press
    sapGui.session.findById("wnd[0]/tbar[0]/btn[3]").press 

    itmnun = int(itmnun) - 1

    #loop para cada item 
    for i in range(int(itmnun + 1)):

        #Recupera storedlocation
        sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl01n"
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").text = str(shipPoint.value)
        sapGui.session.findById("wnd[0]/usr/ctxtLV50C-VBELN").text = str(ovRef.value)
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]").sendVKey(0)

        sapGui.session.findById("wnd[0]").maximize()
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()       

        print(sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPS-POSNR[0," + str(i) +"]").text)

        #posnr =  session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," + str(i) +"]").text
        stLocation = sheet_obj.cell(row = 2, column = 11)
        stLocation.value =  sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-LGORT[3," + str(i) +"]").text

        
    #Delivery save 
    sapGui.session.findById("wnd[0]").sendVKey(11)
    sapGui.session.findById("wnd[1]").sendVKey(0)
    sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    print(sapGui.session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text)
    dlvNum = sheet_obj.cell(row = 2, column = 13)
    dlvNum.value = sapGui.session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text
    wb_obj.save(pathExel)

    #define serial number/ set stock
    sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb1c"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "561"
    sapGui.session.findById("wnd[0]/usr/ctxtRM07M-SOBKZ").text = "E"

    plant = sheet_obj.cell(row = 2, column = 3)

    sapGui.session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = plant.value
    sapGui.session.findById("wnd[0]/usr/ctxtRM07M-GRUND").setFocus
    sapGui.session.findById("wnd[0]/usr/ctxtRM07M-GRUND").caretPosition = 0
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/ctxtMSEGK-MAT_KDAUF").text = ovRef.value

    nrItem = sheet_obj.cell(row = 2, column = 6)
    sapGui.session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/txtMSEGK-MAT_KDPOS").text = nrItem.value

    matnr = sheet_obj.cell(row = 2, column = 2)
    sapGui.session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = matnr.value

    qtdItem = sheet_obj.cell(row = 2, column = 5)
    sapGui.session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = qtdItem.value


    sapGui.session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").text = stLocation.value
    sapGui.session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").setFocus()
    sapGui.session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").caretPosition = 0
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[1]/tbar[0]/btn[7]").press
    sapGui.session.findById("wnd[1]").sendVKey(7)
    serNum = sapGui.session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text
    sapGui.session.findById("wnd[1]").sendVKey(0)
    print(serNum)
    sapGui.session.findById("wnd[0]").sendVKey(11)


    #Picking
    for i in range(int(itmnun + 1)):
        sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").text = qtdItem.value
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").setFocus()
        sapGui.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").caretPosition = 1
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #Set objStdOut = WScript.StdOut
        #objStdOut.Write session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," + str(i) + "]").text
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]").sendVKey(11)
        sapGui.session.findById("wnd[0]").sendVKey(0)

        #set serial number and post good issues
        sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]").resizeWorkingPane(128,37,False)
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/mbar/menu[3]").select()
        sapGui.session.findById("wnd[0]/mbar/menu[3]/menu[3]").select()
        sapGui.session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text = str(serNum)
        sapGui.session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").caretPosition = 3
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]").sendVKey(11)
        sapGui.session.findById("wnd[0]").sendVKey(0)
        sapGui.session.findById("wnd[0]/tbar[1]/btn[20]").press()

    sapGui.session.findById("wnd[0]").resizeWorkingPane(128,37,False)
    sapGui.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]").sendVKey(0)
    sapGui.session.findById("wnd[0]/tbar[0]/btn[11]").press()


    while sapGui.connection.children.count > 0:
        time.sleep(10)
        print("SAP GUI OPEN")


    # session = None
    # connection = None
    # application = None
    # SapGuiAuto = None

SAP_OP()