import win32com.client
import sys
import openpyxl
import time

def SAP_OP():

    excelPath = r'#'
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return

    pathExel = "C:/Automation/Automate_NF_Transicao_Teste.xlsx"
    wb_obj = openpyxl.load_workbook(pathExel) 
    sheet_obj = wb_obj.active 

    ovOrig = sheet_obj.cell(row = 2, column = 1)  

    #script para pegar o numero da OV
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "/USE/PDS3_OBJMAP"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = "PAG"
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "410"
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,3]").text = "SALES_DOC"
    print(ovOrig.value)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = "*13705704*"
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").caretPosition = 1
    #session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = str(ovOrig.value)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarButton("DETAIL")
    session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").setFocus()
    session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").caretPosition = 0
    ovNew = session.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,7]").text

    tpOV = sheet_obj.cell(row = 2, column = 10) 
    ovRef = sheet_obj.cell(row = 2, column = 12)

    #criar OV com referencia
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = tpOV.value
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF").select()
    session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").text = ovNew
    session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").caretPosition = 10
    session.findById("wnd[1]").sendVKey(5)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = " "
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").setFocus
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]").sendVKey(11)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
    session.findById("wnd[0]").sendVKey(0)
    ovRef.value = session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
 
    #Quantidade de itens da OV
    shipPoint = sheet_obj.cell(row = 2, column = 7)
    wb_obj.save(pathExel)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "vbap"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = str(ovRef.value)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").caretPosition = 5
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 3
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 6
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 9
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 12
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 15
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 18
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 21
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 24
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 27
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 30
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 33
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 36
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 39
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 42
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 45
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").text = str(shipPoint.value)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").setFocus
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,25]").caretPosition = 4
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/txtGD-NUMBER").setFocus
    session.findById("wnd[0]/usr/txtGD-NUMBER").caretPosition = 0
    itmnun = session.findById("wnd[0]/usr/txtGD-NUMBER").text
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press 

    itmnun = int(itmnun) - 1

    for i in range(int(itmnun + 1)):

        #get storedlocation
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl01n"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtLIKP-VSTEL").text = str(shipPoint.value)
        session.findById("wnd[0]/usr/ctxtLV50C-VBELN").text = str(ovRef.value)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)

        print(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-LGORT[3," + str(i) +"]").text)
        print(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPS-POSNR[0," + str(i) +"]").text)

        #posnr =  session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," + str(i) +"]").text
        stLocation = sheet_obj.cell(row = 2, column = 11)
        stLocation.value =  session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-LGORT[3," + str(i) +"]").text


        #session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," & %IndexLinha% & "]").setFocus
        #session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," & %IndexLinha% & "]").caretPosition = 0
        #Set objStdOut = WScript.StdOut
        #objStdOut.Write session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," & %IndexLinha% & "]").text

        
    #Delivery save 
    session.findById("wnd[0]").sendVKey(11)
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
    session.findById("wnd[0]").sendVKey(0)
    print(session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text)

    dlvNum = sheet_obj.cell(row = 2, column = 13)
    dlvNum.value = session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text
    wb_obj.save(pathExel)

    #define serial number/ set stock
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb1c"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "561"
    session.findById("wnd[0]/usr/ctxtRM07M-SOBKZ").text = "E"

    plant = sheet_obj.cell(row = 2, column = 3)

    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = plant.value
    session.findById("wnd[0]/usr/ctxtRM07M-GRUND").setFocus
    session.findById("wnd[0]/usr/ctxtRM07M-GRUND").caretPosition = 0
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/ctxtMSEGK-MAT_KDAUF").text = ovRef.value

    nrItem = sheet_obj.cell(row = 2, column = 6)
    session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/txtMSEGK-MAT_KDPOS").text = nrItem.value

    matnr = sheet_obj.cell(row = 2, column = 2)
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = matnr.value

    qtdItem = sheet_obj.cell(row = 2, column = 5)
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = qtdItem.value


    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").text = stLocation.value
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").setFocus
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").caretPosition = 0
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[1]/tbar[0]/btn[7]").press
    session.findById("wnd[1]").sendVKey(7)
    serNum = session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text
    session.findById("wnd[1]").sendVKey(0)
    print(serNum)
    session.findById("wnd[0]").sendVKey(11)


    #Picking
    for i in range(int(itmnun + 1)):
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").text = qtdItem.value
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").setFocus()
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").caretPosition = 1
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #Set objStdOut = WScript.StdOut
        #objStdOut.Write session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," + str(i) + "]").text
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[0]").sendVKey(0)

        #set serial number and post good issues
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").resizeWorkingPane(128,37,False)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/mbar/menu[3]").select()
        session.findById("wnd[0]/mbar/menu[3]/menu[3]").select()
        session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text = str(serNum)
        session.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").caretPosition = 3
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[20]").press()

    session.findById("wnd[0]").resizeWorkingPane(128,37,False)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

    session = None
    connection = None
    application = None
    SapGuiAuto = None

SAP_OP()