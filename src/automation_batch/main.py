import time
import openpyxl
import sys
import os
import tkinter
from tkinter import *
from tkinter import messagebox as mb

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from sku_reader import SKU
import shipstation_api as api

def saveExcel(workbook,workbookPath):#Saves workbook input at workbookPath
    saved = False
    while saved  == False:
        try:
            workbook.save(workbookPath)
            saved = True
        except PermissionError:
            time.sleep(1)
            
def openExcel(path):#Opens Batch.xlsx, the Excel sheet holding all Batches and SKUs
    workbookPath = path
    workbook = openpyxl.load_workbook(workbookPath)
    saveExcel(workbook, workbookPath)
    return workbook, workbookPath

def notesTrim(book, path):
    sheet = book["Notes"]
    row = 1
    orderCell = sheet.cell(row = row, column = 1)
    while orderCell.value != None:
        dateCell = sheet.cell(row = row, column = 2)
        if dateCell.value == None:
            dateCell.value = "NO DATE"
        else:
            dateString = dateCell.value
            dateString = dateString.replace("Ship by ","")
            if '  via ' in dateString:
                dateString = dateString.split("  via ")[0]
            elif ' via ' in dateString:
                dateString = dateString.split(" via ")[0]
            dateCell.value = dateString
            print(str(orderCell.value)+": "+dateString)
        saveExcel(book,path)
        row += 1
        orderCell = sheet.cell(row = row, column = 1)

def notesToBatches(notesBook,batchBook,batchPath):
    notesSheet = notesBook["Notes"]
    notesRow = 2
    notesCell = notesSheet.cell(row = notesRow, column = 1)
    while notesCell.value != None:
        dateCell = str(notesSheet.cell(row = notesRow, column = 2).value)
        if dateCell not in batchBook.sheetnames:
            batchBook.create_sheet(dateCell)
            saveExcel(batchBook,batchPath)
        batchSheet = batchBook[dateCell]
        batchRow = 1
        batchCell = batchSheet.cell(row = batchRow, column = 1)
        while batchCell.value != None and batchCell.value != notesCell.value:
            batchRow += 1
            batchCell = batchSheet.cell(row = batchRow, column = 1)
        if batchCell.value == None:
            batchCell.value = notesCell.value
            print(batchCell.value)
            saveExcel(batchBook,batchPath)
        notesRow += 1
        notesCell = notesSheet.cell(row = notesRow, column = 1)

def orderFill(batchBook, batchPath, ordersBook,ordersPath):
    for batchSheet in batchBook.worksheets:
        print(f"Checking sheet: {batchSheet.title}")
        batchRow = 1
        while True:
            batchCell = batchSheet.cell(row=batchRow, column=1).value
            if batchCell is None:
                break
            ###########################################################
            for orderSheet in ordersBook.worksheets:
                ordersRow = 1
                while True:
                    ordersCell = orderSheet.cell(row=ordersRow, column=1).value
                    if ordersCell is None:
                        break
                    if ordersCell == batchCell:
                        print(f"  FOUND AT: Row {ordersRow}: {ordersCell}")
                        batchSheet.cell(row=batchRow, column=2).value = orderSheet.cell(row=ordersRow, column=2).value#SKU
                        if "Items" in str(orderSheet.cell(row=ordersRow, column=2).value):
                            multi = str(orderSheet.cell(row=ordersRow, column=2).value)
                            multi = multi.replace("(","")
                            multi = multi.replace(" Items)","")
                            batchSheet.cell(row=batchRow, column=2).value = int(multi)
                            print(orderSheet.cell(row=ordersRow, column=2).value)
                            batchSheet.cell(row=batchRow, column=7).value = "Multi"
                        batchSheet.cell(row=batchRow,column=3).value = orderSheet.cell(row=ordersRow,column=3).value#Quantity
                        batchSheet.cell(row=batchRow,column=4).value = orderSheet.cell(row=ordersRow,column=4).value#Batch
                        batchSheet.cell(row=batchRow,column=5).value = orderSheet.cell(row=ordersRow,column=5).value#HoldUntil
                        batchSheet.cell(row=batchRow,column=6).value = orderSheet.cell(row=ordersRow,column=6).value#ShippingRequested
                        batchSheet.cell(row=batchRow,column=11).value = orderSheet.cell(row=ordersRow,column=7).value#ShipFrom
                        saveExcel(batchBook, batchPath)
                        orderSheet.delete_rows(ordersRow)
                        saveExcel(ordersBook,ordersPath)
                        break
                    #print(f"  Row {ordersRow}: {ordersCell}")
                    ordersRow += 1
            ###########################################################
            print(f"  Row {batchRow}: {batchCell}")
            batchRow += 1

    print("Finished checking all sheets.")

def skuCheck(batchBook, batchPath):
    for batchSheet in batchBook.worksheets:
        print(f"Running skuCheck for sheet: {batchSheet.title}")
        batchRow = 1
        while True:
            batchCell = batchSheet.cell(row=batchRow, column=1).value
            if batchCell is None:
                break
            if batchSheet == batchBook['BatchNumber']:
                break
            ###########################################################
            sku_ = batchSheet.cell(row=batchRow, column=2).value
            if batchSheet.cell(row=batchRow, column=7).value != 'Multi':
                sku = SKU(sku_)
                if sku.hd == True: 
                    batchSheet.cell(row = batchRow, column = 7).value = 'HD'
                    batchSheet.cell(row = batchRow, column = 8).value = sku.prodType#16 per batch for watch bands
                    if sku.productmatch == 'HDXAirpod' or sku.productmatch == 'Airpod' or sku.productmatch == 'BudsPro' or sku.productmatch == 'HDXBudsPro':#50 per batch
                        batchSheet.cell(row = batchRow, column = 8).value = 'Airpod'
                        #batchSheet.cell(row = batchRow, column = 9).value = sku.template
                    if sku.productmatch == 'Phone':#14 per batch for 1-2 and 12 per batch for 3-4
                        batchSheet.cell(row = batchRow, column = 9).value = sku.template
                    

                if sku.engraved == True:
                    batchSheet.cell(row = batchRow, column = 7).value = 'Engraved'#16 per batch for watch bands
                    batchSheet.cell(row = batchRow, column = 8).value = sku.prodType#16 per batch for watch bands
                    if sku.productmatch == 'HDXAirpod' or sku.productmatch == 'Airpod' or sku.productmatch == 'BudsPro':#28 for buds, 32 for gen 1 and pro, 30 for gen 3, 
                        batchSheet.cell(row = batchRow, column = 8).value = 'Airpod'
                        #batchSheet.cell(row = batchRow, column = 9).value = sku.template

                if sku.leather == True: 
                    batchSheet.cell(row = batchRow, column = 7).value = 'Leather'#25 per batch
                    #batchSheet.cell(row = batchRow, column = 8).value = sku.prodType#keep 20 and 38 in same batch, same with 22 and 42
                    #if sku.prodType == '38':
                    #    batchSheet.cell(row = batchRow, column = 8).value = '20'
                    ##if sku.prodType == '42':
                    #    batchSheet.cell(row = batchRow, column = 8).value = '22'
                    league = sku.league
                    league = league.replace('/','')
                    #batchSheet.cell(row = batchRow, column = 9).value = league
                    #if sku.brand == True:
                    #    batchSheet.cell(row = batchRow, column = 9).value = 'Brand'

                if sku.steel == True: 
                    batchSheet.cell(row = batchRow, column = 7).value = 'Steel'#25 per batch

                if sku.dyesub == True: 
                    batchSheet.cell(row = batchRow, column = 7).value = 'Dyesub'#25 per batch

                if sku.screenprint == True or sku.stocked == True or sku.blank == True:#25 per batch, raise to 50 or 100 starting on black friday
                    batchSheet.cell(row = batchRow, column = 7).value = 'Screenprint'
                    batchSheet.cell(row = batchRow, column = 8).value = sku.productmatch
                    if sku.productmatch == 'HDXAirpod' or sku.productmatch == 'Airpod' or sku.productmatch == 'BudsPro' or sku.productmatch == 'HDXBudsPro':
                        batchSheet.cell(row = batchRow, column = 8).value = 'Airpods'
                    if sku.productmatch == 'BandCombo':
                        batchSheet.cell(row = batchRow, column = 8).value = 'Band'
                    batchSheet.cell(row = batchRow, column = 9).value = batchSheet.cell(row=batchRow,column=11).value

                if sku.combo == True:
                    batchSheet.cell(row = batchRow, column = 8).value = 'Combo'
                    batchSheet.cell(row = batchRow, column = 9).value = ''

                if sku.theRest == True: batchSheet.cell(row = batchRow, column = 7).value = 'theRest'
                if sku.echo == True: batchSheet.cell(row = batchRow, column = 7).value = 'theRest'
                
                
            batchRow += 1
        saveExcel(batchBook, batchPath)

def batchAssign(batchBook, batchPath, batchNum, listBook, listPath):
    priorityList = [
        'Priority',
        '2 Day', 
        '2-Day', 
        '2nd Day',
        'Expedited', 
        'FedEx2DayOneRate',
        'FedEx 2 Day Guaranteed Shipping',
        'FedEx Home Delivery',
        'FedExStandardOvernight',
        'First Class Package International',
        'Free FedEx 2 Day over $55', 
        'USPS First Class International',
        'USPS Priority Mail 2-3 Day Delivery'
        #'UPS 2nd Day Air' Not actually priority
        ]
    for batchSheet in batchBook.worksheets:
        print(f"Running batchAssign for sheet: {batchSheet.title}")
        listSheet = batchSheet.title
        if listSheet not in listBook.sheetnames:
            listBook.create_sheet(listSheet)
            saveExcel(listBook,listPath)
        listSheet = listBook[listSheet]
        listRow = 1
        while listSheet.cell(row=listRow,column=1).value != None:
            listRow+=1
        #print("Empty listRow found at: "+str(listRow))
        batchRow = 1
        while True:
            batchCell = batchSheet.cell(row=batchRow, column=1).value
            if batchCell is None:
                print('Empty batchCell found at: '+str(batchRow)+' with batchNum: '+str(batchNum))
                break
            if batchSheet == batchBook['BatchNumber']:
                print('BatchNumber sheet break with batchNum: '+str(batchNum))
                break
            ###########################################################
            if batchSheet.cell(row=batchRow,column=4).value == None and batchSheet.cell(row=batchRow,column=5).value == None:
                batchQTY = int(batchSheet.cell(row=batchRow,column=3).value)
                batchShipping = batchSheet.cell(row=batchRow,column=6).value
                if batchSheet.cell(row=batchRow,column=6).value != None:
                    if batchShipping in priorityList:
                        batchSheet.cell(row=batchRow,column=6).value = 'Priority'
                    elif batchShipping not in priorityList:
                        batchSheet.cell(row=batchRow,column=6).value = 'Not Priority'
                    if batchSheet.cell(row=batchRow, column=11).value == 'Walmart' or batchSheet.cell(row=batchRow, column=11).value == 'Target Dropship (UPS)':
                        batchSheet.cell(row=batchRow,column=6).value = 'Priority'
                batchShipping = batchSheet.cell(row=batchRow,column=6).value
                batch1 = batchSheet.cell(row=batchRow,column=7).value
                batch2 = batchSheet.cell(row=batchRow,column=8).value
                batch3 = batchSheet.cell(row=batchRow,column=9).value
                #print('Batch Started: '+str(batchNum)+' at row '+str(batchRow))
                batchSheet.cell(row=batchRow,column=4).value = 'BatchID'+str(batchNum)
                
                listSheet.cell(row=listRow,column=1).value = 'BatchID'+str(batchNum)
                listSheet.cell(row=listRow,column=2).value = batchShipping
                listSheet.cell(row=listRow,column=3).value = batchQTY
                listSheet.cell(row=listRow,column=4).value = batch1
                listSheet.cell(row=listRow,column=5).value = batch2
                listSheet.cell(row=listRow,column=6).value = batch3
                #print('List Row: '+str(listRow)+' and Batch Num: '+str(batchNum))
                saveExcel(listBook, listPath)

                tempRow = batchRow + 1
                while batchSheet.cell(row=tempRow, column=1).value != None:
                    tempQTY = int(batchSheet.cell(row=tempRow,column=3).value)
                    tempShipping = batchSheet.cell(row=tempRow,column=6).value
                    if batchSheet.cell(row=tempRow,column=6).value != None:
                        if tempShipping in priorityList:
                            batchSheet.cell(row=tempRow,column=6).value = 'Priority'
                        else:
                            batchSheet.cell(row=tempRow,column=6).value = 'Not Priority'
                        if batchSheet.cell(row=tempRow, column=11).value == 'Walmart' or batchSheet.cell(row=tempRow, column=11).value == 'Target Dropship (UPS)':
                            batchSheet.cell(row=tempRow,column=6).value = 'Priority'
                    tempShipping = batchSheet.cell(row=tempRow,column=6).value
                    temp1 = batchSheet.cell(row=tempRow,column=7).value
                    temp2 = batchSheet.cell(row=tempRow,column=8).value
                    temp3 = batchSheet.cell(row=tempRow,column=9).value
                    if batchShipping == tempShipping and batch1 == temp1 and batch2 == temp2 and batch3 == temp3:
                        newBatchNum = False
                        checkQTY = batchQTY + tempQTY
                        match temp1:
                            case 'HD':
                                match temp2:
                                    case '20':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '22':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '38':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '42':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case 'Gen1':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'HDXGen1':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'Pro':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'HDXPro':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'Gen3':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'HDXGen3':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'BudsPro':
                                        if checkQTY > 50:
                                            newBatchNum = True
                                    case 'Phone1':
                                        if checkQTY > 14:
                                            newBatchNum = True
                                    case 'Phone2':
                                        if checkQTY > 14:
                                            newBatchNum = True
                                    case 'Phone3':
                                        if checkQTY > 12:
                                            newBatchNum = True
                                    case 'Phone4':
                                        if checkQTY > 12:
                                            newBatchNum = True
                            case 'Engraved':
                                match temp2:
                                    case '20':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '22':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '38':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case '42':
                                        if checkQTY > 16:
                                            newBatchNum = True
                                    case 'Gen1':
                                        if checkQTY > 32:
                                            newBatchNum = True
                                    case 'Pro':
                                        if checkQTY > 32:
                                            newBatchNum = True
                                    case 'Gen3':
                                        if checkQTY > 30:
                                            newBatchNum = True
                                    case 'BudsPro':
                                        if checkQTY > 28:
                                            newBatchNum = True
                            case 'Leather':
                                if checkQTY > 25:
                                    newBatchNum = True
                            case 'Steel':
                                if checkQTY > 25:
                                    newBatchNum = True
                            case 'Dyesub':
                                if checkQTY > 25:
                                    newBatchNum = True
                            case 'Screenprint':
                                if checkQTY > 25:
                                    newBatchNum = True
                            #case 'Multi':
                            #    newBatchNum = False
                            #case 'theRest':
                            #    newBatchNum = False
                        if newBatchNum == True:
                            batchNum += 1
                            listRow += 1
                            listSheet.cell(row=listRow,column=1).value = 'BatchID'+str(batchNum)
                            listSheet.cell(row=listRow,column=2).value = batchShipping
                            listSheet.cell(row=listRow,column=3).value = batchQTY
                            listSheet.cell(row=listRow,column=4).value = batch1
                            listSheet.cell(row=listRow,column=5).value = batch2
                            listSheet.cell(row=listRow,column=6).value = batch3
                            saveExcel(listBook, listPath)
                        batchSheet.cell(row=tempRow,column=4).value = 'BatchID'+str(batchNum)
                        batchQTY = batchQTY + tempQTY
                        if newBatchNum == True:
                            #print('Batch: '+str(batchNum)+' for '+str(batch1)+' '+str(batch2)+' '+str(batch3)+' at row: '+str(tempRow))
                            batchQTY = int(batchSheet.cell(row=tempRow,column=3).value)
                        else:
                            listSheet.cell(row=listRow,column=3).value = batchQTY
                    tempRow+=1
                #print('Batch: '+str(batchNum)+' for '+str(batch1)+' '+str(batch2)+' '+str(batch3))
                listRow+=1
                batchNum+=1
            batchRow+=1

        saveExcel(batchBook, batchPath)
    return batchNum

def fAPI(batchNum,orderList,listBook,listPath):
    batchId, shstId = api.makeEmptyBatch(batchNum)#shstId is the id on the front end of ship station
    for listSheet in listBook.worksheets:
        listRow = 1
        while listSheet.cell(row=listRow,column=1).value != None:
            if listSheet.cell(row=listRow,column=1).value == batchNum:
                listSheet.cell(row=listRow,column=1).value = shstId
                saveExcel(listBook, listPath)
            listRow += 1
    for order in orderList:
        shipmentId = api.getOrder(order)
        #rateId = api.getRateId(shipmentId,carrier,shipment,package)
        api.addToBatch(batchId,shipmentId)

def getOrders(batchBook,batchPath,listBook,listPath):
    for batchSheet in batchBook.worksheets:
        print(f"Running getOrder for sheet: {batchSheet.title}")
        batchRow = 1
        batchOld = ''
        while True:
            batchCell = batchSheet.cell(row=batchRow, column=1).value
            if batchCell is None:
                print('Empty batchCell found at: '+str(batchRow))
                break
            if batchSheet == batchBook['BatchNumber']:
                print('BatchNumber sheet break')
                break
            ###########################################################
            if batchSheet.cell(row=batchRow,column=4).value != None:# and batchSheet.cell(row=batchRow,column=10).value != 'BatchAssigned':#Column J for batch assigned
                batchNum = batchSheet.cell(row=batchRow,column=4).value
                if batchOld == '':#Start batchList[]
                    batchList = []
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                elif batchNum == batchOld:#Append to batchList[]
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                elif batchOld != batchNum:#New batchList[]
                    print('Batch: '+str(batchOld)+': '+str(batchList))
                    fAPI(batchOld,batchList,listBook,listPath)
                    #batchSheet.cell(row=batchRow,column=10).value = 'BatchAssigned'
                    saveExcel(batchBook, batchPath)
                    batchList = []
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                if batchSheet.cell(row=batchRow+1,column=4).value == None:
                    print('Batch: '+str(batchNum)+': '+str(batchList))
                    fAPI(batchNum,batchList,listBook,listPath)
                    #batchSheet.cell(row=batchRow,column=10).value = 'BatchAssigned'
                    saveExcel(batchBook, batchPath)
                batchOld = batchNum
            batchRow+=1
                
def main():
    batchBook, batchPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Batches.xlsx")
    notesBook, notesPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Notes.xlsx")
    ordersBook, ordersPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Orders.xlsx")
    listBook, listPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/BatchList.xlsx")
    
    batchNumSheet = batchBook["BatchNumber"]
    batchNum = int(batchNumSheet.cell(row=1,column=1).value)
    
    notesTrim(notesBook,notesPath)
    notesToBatches(notesBook,batchBook,batchPath)
    orderFill(batchBook, batchPath, ordersBook, ordersPath)
    skuCheck(batchBook, batchPath)
    #tkinter.messagebox.showinfo(title='BatchAssign', message='Check shipping in Batches.xlsx')
    batchNum = batchAssign(batchBook, batchPath, batchNum, listBook, listPath)
    batchNumSheet.cell(row=1,column=1).value = batchNum
    saveExcel(batchBook, batchPath)

    tkinter.messagebox.showinfo(title='Sort', message='Manually sort Batches.xlsx')

    batchBook, batchPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Batches.xlsx")
    listBook, listPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/BatchList.xlsx")

    #Start the Batch Making process
    getOrders(batchBook,batchPath,listBook,listPath)

    print("Done")

main()