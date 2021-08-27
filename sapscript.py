#import win32com.client
import sys
import openpyxl
import time
import loginSAP
from   runScripstClass import RunScripts
import constants

def SAP_OP():

    #call class open SAP GUI
    sapGui  =  loginSAP.SapGui()

    pathExel = constants.FILENAME
    wb_obj = openpyxl.load_workbook(pathExel) 
    sheet_obj = wb_obj.active
    ovOrig = sheet_obj.cell(row = 2, column = 1) 
    
    runScript = RunScripts(sapGui)

    #Get SO from PRD
    try: 
        ovNew = runScript.getSOfromProd(ovOrig.value)

    except:
        print('Falha ao recuperar doc de PRD')
        print("next document")


    #get from spreadsheet Type SO and set SO Ref
    tpOV = sheet_obj.cell(row = 2, column = 10) 
    ovRef = sheet_obj.cell(row = 2, column = 12)

    #Create SO with reference
    try:

        ovRef.value = runScript.createSOByRef(tpOV.value, ovNew)

    except:
        print(sys.exc_info()[0])
    
 
    #Save Shiping point into Spredsheet
    shipPoint = sheet_obj.cell(row = 2, column = 7)
    wb_obj.save(pathExel)

    #get ITM NUM by SO
    try:

        itmnun = runScript.getItmNum(ovRef.value, shipPoint.value)

    except:
        print(sys.exc_info()[0])

    itmnun = int(itmnun) - 1

    #get Stored location by ITMUN
    for i in range(int(itmnun + 1)):

        try:

            stLocation = sheet_obj.cell(row = 2, column = 11)
            stLocation.value = runScript.getStoredLoc(shipPoint.value, ovRef.value, i)

        except:
            print(sys.exc_info()[0])


    #Delivery save 
    try:

        deliveryValue = runScript.getDelivery()

    except:
        print(sys.exc_info()[0])


    #Save Delivery into SpreadSheet
    dlvNum = sheet_obj.cell(row = 2, column = 13)
    dlvNum2 = sheet_obj.cell(row = 2, column = 14)

    dlvNum.value = deliveryValue
    dlvNum2.value = deliveryValue
    wb_obj.save(pathExel)


    #define serial number/ set stock
    plant = sheet_obj.cell(row = 2, column = 3)
    nrItem = sheet_obj.cell(row = 2, column = 6)
    matnr = sheet_obj.cell(row = 2, column = 2)
    qtdItem = sheet_obj.cell(row = 2, column = 5)

    try:
        serNum = sheet_obj.cell(row = 2, column = 8)
        serNum.value = runScript.setStockAndSerNumber(ovRef.value, plant.value, nrItem.value, matnr.value, qtdItem.value, stLocation.value)

    except:
        print(sys.exc_info()[0])

    
    #SAVE SERIAL NUMBER
    wb_obj.save(pathExel)


    #Picking
    for i in range(int(itmnun + 1)):

        try:
            #Picking
            runScript.setDefPicking(qtdItem.value)

            #set serial number and post good issues
            runScript.setSerialNumAndPostGI(serNum.value)

            #Set Serial number at DLV

        except:
            print(sys.exc_info()[0])


    #get billing doc
    billDoc = sheet_obj.cell(row = 2, column = 15)
    billDoc.value = runScript.createBilling()
    wb_obj.save(pathExel)

    #GET GOODS MVT DOC
    goodMov = sheet_obj.cell(row = 2, column = 16)
    goodMov.value = runScript.getGoodsMvt(ovRef.value, nrItem.value, constants.GOOD_MOV)    
    wb_obj.save(pathExel)

    while sapGui.connection.children.count > 0:
        time.sleep(10)
        print("SAP GUI OPEN")


    # session = None
    # connection = None
    # application = None
    # SapGuiAuto = None

SAP_OP()