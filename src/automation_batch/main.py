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
    notesRow = 1
    orderCell = sheet.cell(row=notesRow,column=1)
    while orderCell.value != None:
        dateCell = sheet.cell(row=notesRow,column=2)
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
        notesRow += 1
        orderCell = sheet.cell(row=notesRow,column=1)

def notesToBatches(notesBook,batchBook,batchPath):
    notesSheet = notesBook["Notes"]
    notesRow = 2
    notesCell = notesSheet.cell(row=notesRow,column=1)
    while notesCell.value != None:
        dateCell = str(notesSheet.cell(row=notesRow,column=2).value)
        if dateCell not in batchBook.sheetnames:
            batchBook.create_sheet(dateCell)
            saveExcel(batchBook,batchPath)
        batchSheet = batchBook[dateCell]
        batchRow = 1
        batchCell = batchSheet.cell(row=batchRow,column=1)
        while batchCell.value != None and batchCell.value != notesCell.value:
            batchRow += 1
            batchCell = batchSheet.cell(row=batchRow,column=1)
        if batchCell.value == None:
            batchCell.value = notesCell.value
            print(batchCell.value)
            saveExcel(batchBook,batchPath)
        notesRow += 1
        notesCell = notesSheet.cell(row=notesRow,column=1)

def orderFill(batchBook, batchPath, ordersBook,ordersPath):
    for batchSheet in batchBook.worksheets:
        print(f"Checking sheet: {batchSheet.title}")
        batchRow = 1
        while True:
            batchCell = batchSheet.cell(row=batchRow,column=1).value
            if batchCell is None:
                break
            ###########################################################
            for orderSheet in ordersBook.worksheets:
                ordersRow = 1
                while True:
                    ordersCell = orderSheet.cell(row=ordersRow,column=1).value
                    if ordersCell is None:
                        break
                    if ordersCell == batchCell:
                        print(f"  FOUND AT: Row {ordersRow}: {ordersCell}")
                        batchSheet.cell(row=batchRow,column=2).value = orderSheet.cell(row=ordersRow,column=2).value#SKU
                        if "Items" in str(orderSheet.cell(row=ordersRow,column=2).value):
                            multi = str(orderSheet.cell(row=ordersRow,column=2).value)
                            multi = multi.replace("(","")
                            multi = multi.replace(" Items)","")
                            batchSheet.cell(row=batchRow,column=2).value = int(multi)
                            print(orderSheet.cell(row=ordersRow,column=2).value)
                            batchSheet.cell(row=batchRow,column=8).value = "Multi"
                        batchSheet.cell(row=batchRow,column=3).value = orderSheet.cell(row=ordersRow,column=3).value#Quantity
                        batchSheet.cell(row=batchRow,column=4).value = orderSheet.cell(row=ordersRow,column=4).value#Batch
                        batchSheet.cell(row=batchRow,column=5).value = orderSheet.cell(row=ordersRow,column=5).value#HoldUntil
                        batchSheet.cell(row=batchRow,column=6).value = orderSheet.cell(row=ordersRow,column=6).value#ShippingRequested
                        batchSheet.cell(row=batchRow,column=7).value = orderSheet.cell(row=ordersRow,column=7).value#ShipFrom
                        saveExcel(batchBook, batchPath)
                        orderSheet.delete_rows(ordersRow)
                        saveExcel(ordersBook,ordersPath)
                        break
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
            batchCell = batchSheet.cell(row=batchRow,column=1).value
            if batchCell is None:
                break
            if batchSheet == batchBook['BatchNumber']:
                break
            ###########################################################
            sku_ = batchSheet.cell(row=batchRow,column=2).value
            if batchSheet.cell(row=batchRow,column=8).value != 'Multi':
                sku = SKU(sku_)

                if sku.hd == True: 
                    batchSheet.cell(row=batchRow,column=8).value = 'HD'
                    batchSheet.cell(row=batchRow,column=8).value = sku.prodType
                    if batchSheet.cell(row=batchRow,column=8).value == '38':
                        batchSheet.cell(row=batchRow,column=8).value = 'HD 38 L+S'
                    if batchSheet.cell(row=batchRow,column=8).value == '42':
                        batchSheet.cell(row=batchRow,column=8).value = 'HD 42 L+S'
                    if batchSheet.cell(row=batchRow,column=8).value == '20' or batchSheet.cell(row=batchRow,column=8).value == '22':
                        batchSheet.cell(row=batchRow,column=8).value = 'HD 20+22'
                    if sku.productmatch == 'Airpod' or sku.productmatch == 'BudsPro' or sku.productmatch == 'HDXAirpod' or sku.productmatch == 'HDXBudsPro':
                        batchSheet.cell(row=batchRow,column=8).value = 'HD Airpods'
                    if sku.productmatch == 'Phone':
                        batchSheet.cell(row=batchRow,column=8).value = 'Phones'
                        if sku.phone == 'AA17' or sku.phone == 'AP17' or sku.phone == 'AM17' or sku.pixelphone == True:
                            batchSheet.cell(row=batchRow,column=8).value = 'iPhone 17'

                if sku.engraved == True or sku.leather == True:
                    batchSheet.cell(row=batchRow,column=8).value = 'Engraved+Leather'

                if sku.steel == True or sku.dyesub == True: 
                    batchSheet.cell(row=batchRow,column=8).value = 'Steel+Dyesub'#25 per batch

                if sku.screenprint == True or sku.stocked == True or sku.blank == True:#25 per batch, raise to 50 or 100 starting on black friday
                    batchSheet.cell(row=batchRow,column=8).value = 'Screenprint'

                if sku.combo == True:
                    batchSheet.cell(row=batchRow,column=8).value = 'Multi'

                if sku.theRest == True:
                    batchSheet.cell(row=batchRow,column=8).value = 'theRest'
                if sku.echo == True:
                    batchSheet.cell(row=batchRow,column=8).value = 'theRest'
                
            batchRow += 1
        saveExcel(batchBook, batchPath)

def batchAssign(batchBook, batchPath, batchNum):
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
        'Free FedEx 2 Day over $75',
        'USPS First Class International',
        'USPS Priority Mail 2-3 Day Delivery'
        #'UPS 2nd Day Air' Not actually priority
        ]
    for batchSheet in batchBook.worksheets:
        print(f"Running batchAssign for sheet: {batchSheet.title}")
        batchRow = 1
        while True:
            batchCell = batchSheet.cell(row=batchRow,column=1).value
            if batchCell is None:
                print('Empty batchCell found at: '+str(batchRow)+' with batchNum: '+str(batchNum))
                break
            if batchSheet == batchBook['BatchNumber']:
                print('BatchNumber sheet break with batchNum: '+str(batchNum))
                break
            ###########################################################
            if batchSheet.cell(row=batchRow,column=4).value == None and batchSheet.cell(row=batchRow,column=5).value == None:
                batchShipping = batchSheet.cell(row=batchRow,column=6).value
                if batchSheet.cell(row=batchRow,column=6).value != None:
                    if batchShipping in priorityList:
                        batchSheet.cell(row=batchRow,column=6).value = 'Priority'
                    elif batchShipping not in priorityList:
                        batchSheet.cell(row=batchRow,column=6).value = 'Not Priority'
                    if batchSheet.cell(row=batchRow,column=7).value == 'Walmart' or batchSheet.cell(row=batchRow,column=7).value == 'Target Dropship (UPS)':
                        batchSheet.cell(row=batchRow,column=6).value = 'Priority'
                batchShipping = batchSheet.cell(row=batchRow,column=6).value
                batch1 = batchSheet.cell(row=batchRow,column=8).value
                batchSheet.cell(row=batchRow,column=4).value = 'BatchID'+str(batchNum)

                tempRow = batchRow + 1
                while batchSheet.cell(row=tempRow,column=1).value != None:
                    tempShipping = batchSheet.cell(row=tempRow,column=6).value
                    if batchSheet.cell(row=tempRow,column=6).value != None:
                        if tempShipping in priorityList:
                            batchSheet.cell(row=tempRow,column=6).value = 'Priority'
                        else:
                            batchSheet.cell(row=tempRow,column=6).value = 'Not Priority'
                        if batchSheet.cell(row=tempRow,column=7).value == 'Walmart' or batchSheet.cell(row=tempRow,column=7).value == 'Target Dropship (UPS)':
                            batchSheet.cell(row=tempRow,column=6).value = 'Priority'
                    tempShipping = batchSheet.cell(row=tempRow,column=6).value
                    temp1 = batchSheet.cell(row=tempRow,column=8).value
                    
                    if batchShipping == 'Priority' and batchShipping == tempShipping:
                        batchSheet.cell(row=tempRow,column=4).value = 'BatchID'+str(batchNum)
                    elif batch1 == temp1 and batchShipping == tempShipping:
                        batchSheet.cell(row=tempRow,column=4).value = 'BatchID'+str(batchNum)

                    tempRow+=1
                batchNum+=1
            batchRow+=1

        saveExcel(batchBook, batchPath)
    return batchNum

def fAPI(batchNum,orderList,batchNotes):
    batchId = api.makeEmptyBatch(batchNum,batchNotes)#shstId is the id on the front end of ship station
    for order in orderList:
        shipmentId = api.getOrder(order)
        #rateId = api.getRateId(shipmentId,carrier,shipment,package)
        api.addToBatch(batchId,shipmentId)

def getOrders(batchBook,batchPath):
    for batchSheet in batchBook.worksheets:
        print(f"Running getOrder for sheet: {batchSheet.title}")
        batchRow = 1
        batchOld = ''
        batchShipping = ''
        batchType = ''
        while True:
            batchCell = batchSheet.cell(row=batchRow,column=1).value
            if batchCell is None:
                print('Empty batchCell found at: '+str(batchRow))
                break
            if batchSheet == batchBook['BatchNumber']:
                print('BatchNumber sheet break')
                break
            ###########################################################
            if batchSheet.cell(row=batchRow,column=4).value != None:
                batchNum = batchSheet.cell(row=batchRow,column=4).value
                #batchShipping = ''
                #batchType = ''

                if batchOld == '':#Start batchList[]
                    batchList = []
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                    batchShipping = batchSheet.cell(row=batchRow,column=6).value
                    batchType = str(batchSheet.cell(row=batchRow,column=8).value)
                    #print('first batch type: '+batchType)

                elif batchNum == batchOld:#Append to batchList[]
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                    #batchShipping = batchSheet.cell(row=batchRow,column=6).value
                    #batchType = str(batchSheet.cell(row=batchRow,column=8).value)
                    #print('append batch type: '+batchType)

                elif batchOld != batchNum:#New batchList[]
                    print('Batch: '+str(batchOld)+': '+str(batchList))
                    if batchShipping == 'Priority':
                        batchNotes = 'Priority'
                        #print('priority batch type: '+batchType)
                        print(batchNotes)
                    else:
                        #print('non-prio batch type: '+batchType)
                        batchNotes = batchType
                        print(batchNotes)
                    fAPI(batchOld,batchList,batchNotes)
                    saveExcel(batchBook, batchPath)
                    batchList = []
                    batchList.append(batchSheet.cell(row=batchRow,column=1).value)
                    batchShipping = batchSheet.cell(row=batchRow,column=6).value
                    batchType = str(batchSheet.cell(row=batchRow,column=8).value)

                if batchSheet.cell(row=batchRow+1,column=4).value == None:
                    print('!!!!!!!!!!!FINAL BATCH!!!!!!!!!!')
                    print('Batch: '+str(batchNum)+': '+str(batchList))
                    if batchShipping == 'Priority':
                        batchNotes = 'Priority'
                        print('priority batch type: '+batchType)
                        print(batchNotes)
                    else:
                        print('non-prio batch type: '+batchType)
                        batchNotes = batchType
                        print(batchNotes)
                    fAPI(batchNum,batchList,batchNotes)
                    saveExcel(batchBook, batchPath)
                batchOld = batchNum
            batchRow+=1
                
def main():
    batchBook, batchPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Batches.xlsx")
    notesBook, notesPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Notes.xlsx")
    ordersBook, ordersPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Orders.xlsx")
    
    batchNumSheet = batchBook["BatchNumber"]
    batchNum = int(batchNumSheet.cell(row=1,column=1).value)
    
    tk = Tk()
    tk.geometry('100x150')
    tk.eval('tk::PlaceWindow . center')
    var1 = IntVar()
    Checkbutton(tk, text='Excel Sheets', variable=var1).grid(row=0, sticky = W)
    var2 = IntVar()
    Checkbutton(tk, text='API Calls', variable=var2).grid(row=1, sticky = W)
    Button(tk, text='Run', command=tk.destroy).grid(row=2, sticky = N)
    tk.mainloop()
    if var1.get() == 1:
        notesTrim(notesBook,notesPath)
        notesToBatches(notesBook,batchBook,batchPath)
        orderFill(batchBook, batchPath, ordersBook, ordersPath)
        skuCheck(batchBook, batchPath)
        batchNum = batchAssign(batchBook, batchPath, batchNum)
        batchNumSheet.cell(row=1,column=1).value = batchNum
        saveExcel(batchBook, batchPath)
        tkinter.messagebox.showinfo(title='Sort', message='Manually sort Batches.xlsx')
    if var2.get() == 1:
        batchBook, batchPath = openExcel("S:/Workstation - DavidK/Code/Automation_Batching/Batches.xlsx")

        #Start the Batch Making process
        getOrders(batchBook,batchPath)

    print("Done")

main()