#! Python 3
# - Copy and Paste Ranges using OpenPyXl library
import os
import openpyxl
import datetime
#Prepare the spreadsheets to copy from and paste.

#File to be copied
wb = openpyxl.load_workbook("Sample1.xlsx") #Add file name
copySheet = wb["Area 02"] #Add Sheet name

# #File to be pasted into
template = openpyxl.load_workbook("Sample2.xlsx") #Add file name
pasteSheet = template["Area 02"] #Add Sheet name
# print(pasteSheet.max_row )

#this returns the column place for the main headers ex. tag in column 2 and problem in column 3
def searchForWord(sheet, theWord):
    for i in range(1, 10,1):
        #Appends the row to a RowSelected list
        for j in range(1, 10,1):
            if theWord in str(sheet.cell(i,j).value).upper():
                return i,j

tagPaste= searchForWord(pasteSheet, "TAG")[1]
tagCopy= searchForWord(copySheet, "TAG")[1]

problemPaste= searchForWord(pasteSheet,"PROBLEM")[1]
problemCopy= searchForWord(copySheet,"PROBLEM")[1]

complainPaste= searchForWord(pasteSheet,"COMP")[1]
complainCopy= searchForWord(copySheet,"COMP")[1]

actionPaste= searchForWord(pasteSheet,"ACTION")[1]
actionCopy= searchForWord(copySheet,"ACTION")[1]

statusPaste= searchForWord(pasteSheet,"STATUS")[1]
statusCopy= searchForWord(copySheet,"STATUS")[1]

datePaste= searchForWord(pasteSheet,"DATE")[1]
dateCopy= searchForWord(copySheet,"DATE")[1]

# A dictionary for the label tags with a column number for each tag.
# later need a row iteration to loop through data.
#ex tag:2 problem:3 Complain:4
nonFilteredPasteDict= {
 "Tag": tagPaste, "Problem": problemPaste,
 "Complain": complainPaste , "Action" : actionPaste, 
 "Status": statusPaste, "Date": datePaste
 }

nonFilterdedCopyDict= {
 "Tag": tagCopy, "Problem": problemCopy,
 "Complain": complainCopy , "Action" : actionCopy, 
 "Status": statusCopy, "Date": dateCopy
 }

# drop down any key value pair if empty value
def dictValidator(d):
    new={}
    for k,v in d.items():
        if v!= None:
            new[k]=v
    return new

copyDict= dictValidator(nonFilterdedCopyDict)
pasteDict= dictValidator(nonFilteredPasteDict)

#Copy range of cells as a nested list
#Takes: start , and sheet you want to copy from.
def copyRange(copyDict, sheet):
    startRow = searchForWord(sheet, "TAG")[0] + 1
    rangeSelected = []
    #Loops through selected Rows.
    while str( sheet.cell(startRow, copyDict["Tag"]).value)!= "None":
        
    # for i in range(12,14,1):
        #Appends the row to a RowSelected list
        rowSelected = {}
        for j in copyDict:
            if isinstance( sheet.cell(row= startRow , column= copyDict[j]).value , datetime.datetime):
                d =sheet.cell( row= startRow , column= copyDict[j]).value
                rowSelected[j]  = d.strftime("%d/%m/%Y")
            else:
                rowSelected[j] =sheet.cell(row= startRow , column= copyDict[j]).value
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
        startRow= startRow + 1
    return rangeSelected


#Paste data from copyRange into template sheet


def pasteRange(copyDict, pasteDict, sheetReceiving, copiedData):
    countRow = 0
    startRow= sheetReceiving.max_row +1
    #Check last row that it is not empty
    while str(sheetReceiving.cell(startRow-1, pasteDict["Tag"] ).value)== "None":
        startRow= startRow-1         #decrement the row bec data was manualyy deleted
    # print(startRow, endRow)
    endRow= startRow+ len(copiedData)
    for i in range(startRow, endRow,1):
        for j in copyDict:
            sheetReceiving.cell(i,pasteDict[j] ).value = copiedData[countRow][j]
        countRow += 1

def createData():
    print("Processing...")
    selectedRange = copyRange(copyDict, copySheet)
    pasteRange(copyDict, pasteDict, pasteSheet, selectedRange) 
    #You can save the template as another file to create a new file here too
    template.save("Sample2.xlsx")
    print("items copied and pasted!")
    return selectedRange

createData()


# open all the files in a folder to copy from
pasteFile= input("File to paste into:")

copyFolder= input("Folder to be copied: ")
print(copyFolder)
os.chdir(copyFolder)
f= os.listdir(copyFolder)
for i in f:
    if i.endswith('.xlsx'):
        createData(i)
