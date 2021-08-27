import subprocess
import time
import win32com.client
import sys

class RunScripts(object):


    def __init__(self, conn):
        
        self.SapGuiSession = conn.session

    def getSOfromProd(self, ovOrig):

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

    def createSOByRef(self, tpSO, SOProd):
        try:

            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtVBAK-AUART").text = tpSO
            self.SapGuiSession.findById("wnd[0]").sendVKey(8)
            self.SapGuiSession.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF").select()
            self.SapGuiSession.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").text = SOProd
            self.SapGuiSession.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRAUF/ssubSUB1:SAPLV45C:0301/ctxtLV45C-VBELN").caretPosition = 10
            self.SapGuiSession.findById("wnd[1]").sendVKey(5)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = " "
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").setFocus
            self.SapGuiSession.findById("wnd[0]/tbar[0]/btn[11]").press
            self.SapGuiSession.findById("wnd[0]").sendVKey(11)
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            return  str(self.SapGuiSession.findById("wnd[0]/usr/ctxtVBAK-VBELN").text)
        
        except:
            
            raise print(sys.exc_info()[0])

    def getItmNum(self, sORef, shipPoint):
        try:
            
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtGD-TAB").text = "vbap"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = str(sORef)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").caretPosition = 10
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/txtGS_SELFIELDS-FIELDNAME[6,1]").caretPosition = 5
            
            self.SapGuiSession.findById("wnd[0]/tbar[0]/btn[71]").press()
            self.SapGuiSession.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = "VSTEL"
            self.SapGuiSession.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
            self.SapGuiSession.findById("wnd[1]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]").text = str(shipPoint)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,0]").caretPosition = 4
            self.SapGuiSession.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            self.SapGuiSession.findById("wnd[0]/usr/txtGD-NUMBER").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/txtGD-NUMBER").caretPosition = 0
            return str(self.SapGuiSession.findById("wnd[0]/usr/txtGD-NUMBER").text)
        
        except:
            
            raise print(sys.exc_info()[0])

    def getStoredLoc(self, shipPoint, sORef, index):
        try:
            
            #Recupera storedlocation
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvl01n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtLIKP-VSTEL").text = str(shipPoint)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtLV50C-VBELN").text = str(sORef)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()

            stLocation = str(self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-LGORT[3," + str(index) +"]").text)


            self.SapGuiSession.findById("wnd[0]").sendVKey(11)
            self.SapGuiSession.findById("wnd[1]").sendVKey(0)      

            return stLocation
            
        except:
            
            raise print(sys.exc_info()[0])

    def getDelivery(self):
        try:
            
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)

            return str(self.SapGuiSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").text)
                    
        except:
            
            raise print(sys.exc_info()[0])

    def setStockAndSerNumber(self, sOref, plant, nrItem, matnr, qtdItem, stLocation):
        try:
            
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nmb1c"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "561"
            self.SapGuiSession.findById("wnd[0]/usr/ctxtRM07M-SOBKZ").text = "E"
            self.SapGuiSession.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = plant
            self.SapGuiSession.findById("wnd[0]/usr/ctxtRM07M-GRUND").setFocus
            self.SapGuiSession.findById("wnd[0]/usr/ctxtRM07M-GRUND").caretPosition = 0
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/ctxtMSEGK-MAT_KDAUF").text = sOref
            self.SapGuiSession.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2423/txtMSEGK-MAT_KDPOS").text = nrItem
            self.SapGuiSession.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = matnr
            self.SapGuiSession.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = qtdItem

            self.SapGuiSession.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").text = stLocation
            self.SapGuiSession.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").caretPosition = 0
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[1]/tbar[0]/btn[7]").press
            self.SapGuiSession.findById("wnd[1]").sendVKey(7)
            serNumber = str(self.SapGuiSession.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text)
            self.SapGuiSession.findById("wnd[1]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(11)

            return serNumber
                    
        except:
            
            raise print(sys.exc_info()[0])

    def setDefPicking(self, qtd):

        try:
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02").select()
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").text = qtd
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").setFocus()
            self.SapGuiSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6,0]").caretPosition = 1
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[1]/tbar[0]/btn[0]").press()

            #Set objStdOut = WScript.StdOut
            #objStdOut.Write session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0," + str(i) + "]").text
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(11)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)

            return

        except:
            
            raise print(sys.exc_info()[0])

    def setSerialNumAndPostGI(self, serNum):

        try:
            #set serial number and post good issues
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvl02n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").resizeWorkingPane(128,37,False)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/mbar/menu[3]").select()
            self.SapGuiSession.findById("wnd[0]/mbar/menu[3]/menu[3]").select()
            self.SapGuiSession.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").text = str(serNum)
            self.SapGuiSession.findById("wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS/ctxtRIPW0-SERNR[0,0]").caretPosition = 3
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(11)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/tbar[1]/btn[20]").press()

            return

        except:
            
            raise print(sys.exc_info()[0])

    def createBilling(self):

        try:
            self.SapGuiSession.findById("wnd[0]").resizeWorkingPane(128,37,False)
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nvf02"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)

            return str(self.SapGuiSession.findById("wnd[0]/usr/ctxtVBRK-VBELN").text)

        except:
            
            raise print(sys.exc_info()[0])

    def getGoodsMvt(self, soNew, nrItem, tpDoc):

        try:

            self.SapGuiSession.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16n"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/ctxtGD-TAB").text = "VBFA"
            self.SapGuiSession.findById("wnd[0]").sendVKey(0)
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = soNew
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = nrItem
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,5]").text = tpDoc
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,5]").setFocus
            self.SapGuiSession.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,5]").caretPosition = 1
            self.SapGuiSession.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.SapGuiSession.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").doubleClickCurrentCell()

            gvmDoc = str(self.SapGuiSession.findById("wnd[1]/usr/tblSAPLTSWUSLDETAIL_TAB/txtGT_SELFIELDS-LOW[1,3]").text)

            self.SapGuiSession.findById("wnd[1]").close()
            self.SapGuiSession.findById("wnd[0]/tbar[0]/btn[15]").press()
            self.SapGuiSession.findById("wnd[0]/tbar[0]/btn[15]").press()

            return gvmDoc

        except:
            
            raise print(sys.exc_info()[0])