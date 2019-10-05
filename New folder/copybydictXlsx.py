#! Python 3
# - Copy and Paste Ranges using OpenPyXl library
import openpyxl ,datetime, os

#Copy range of cells as a nested list
#Takes: start , and sheet you want to copy from.
def copyRange(copyDict, sheet):
    startRow = searchForWord(sheet, "TAG")[0] + 1
    rangeSelected = []
    #Loops through selected Rows.
    while str( sheet.cell(startRow, copyDict["Tag"]).value)!= "None": #the while is to loop until data is finished
        rowSelected = {}
        for j in copyDict:
            rowSelected[j] =sheet.cell(row= startRow, column= copyDict[j]).value
            # if isinstance( sheet.cell(row= startRow , column= copyDict[j]).value , datetime.datetime):
            #     #here we check the type of data copied if date? it must be formatted as must.
            #     d =sheet.cell( row= startRow , column= copyDict[j]).value
            #     rowSelected[j]  = d.strftime("%d/%m/%Y")
            # else:
            #     rowSelected[j] =sheet.cell(row= startRow , column= copyDict[j]).value
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
        startRow= startRow + 1
    return rangeSelected


#Paste data from copyRange into template sheet
def pasteRange(copyDict, pasteDict, sheetReceiving, copiedData, datePaste):
    countRow = 0
    startRow= sheetReceiving.max_row +1
    #Check last row that it is not empty
    while str(sheetReceiving.cell(startRow-1, pasteDict["Tag"] ).value)== "None":
        startRow= startRow-1         #decrement the row to check if the data was manualy deleted

    endRow= startRow+ len(copiedData)
    # pasting is in here:
    for i in range(startRow, endRow,1):  # for every row, we start to paste the new row
        try:  #pasting a new serial number
            sheetReceiving.cell(i,1).value = sheetReceiving.cell(i-1,1).value +1
        except:
            sheetReceiving.cell(i,1).value= 1
        sheetReceiving.cell(i,pasteDict["Date"]).value = datePaste    # pasting date
        

        for j in copyDict:   #for every column, we paste matched tags
            sheetReceiving.cell(i,pasteDict[j] ).value = copiedData[countRow][j]
        countRow += 1

#this returns the column place for the main headers ex. tag in column 2 and problem in column 3
def searchForWord(sheet, theWord):
    for i in range(1, 10,1):
        #Appends the row to a RowSelected list
        for j in range(1, 10,1):
            if theWord in str(sheet.cell(i,j).value).upper():
                return i,j

# drop down any key value pair if empty value
def dictValidator(d):
    new={}
    for k,v in d.items():
        if v!= None:
            new[k]=v
    return new

def createData(copyFileName, pasteFileName):
    print("Processing...")
    copyFile = openpyxl.load_workbook(copyFileName) 
    copySheet= copyFile["Area 04"]

    pasteFile = openpyxl.load_workbook(pasteFileName)
    pasteSheet= pasteFile["Area 02"]

    print("Files are successfully loaded!")

    #To extract the date value from the file name
    dateCopy = os.path.basename(copyFileName).split('.')[0]
    if "," in dateCopy:
        dateCopy= copyFileName.split(",")[0] +"-"+ copyFileName.split("-",maxsplit=1)[1]  
    dateCopy= str(dateCopy)

    tagPaste= searchForWord(pasteSheet, "TAG")[1]
    tagCopy= searchForWord(copySheet, "TAG")[1]

    problemPaste= searchForWord(pasteSheet,"PROB")[1]
    problemCopy= searchForWord(copySheet,"PROB")[1]

    complainPaste= searchForWord(pasteSheet,"COMP")[1]
    complainCopy= searchForWord(copySheet,"COMP")[1]

    actionPaste= searchForWord(pasteSheet,"ACTION")[1]
    actionCopy= searchForWord(copySheet,"ACTION")[1]

    statusPaste= searchForWord(pasteSheet,"STATUS")[1]
    statusCopy= searchForWord(copySheet,"STATUS")[1]

    datePaste= searchForWord(pasteSheet,"DATE")[1]

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
    "Status": statusCopy
    }

    copyDict= dictValidator(nonFilterdedCopyDict)
    pasteDict= dictValidator(nonFilteredPasteDict)

    selectedRange = copyRange(copyDict, copySheet)
    pasteRange(copyDict, pasteDict, pasteSheet, selectedRange, dateCopy) 
    #You can save the template as another file to create a new file here too
    pasteFile.save("Sample2.xlsx")
    print(pasteFile, pasteFileName)
    print("items copied and pasted!")
    return selectedRange

# createData("Sample1.xlsx", "Sample2.xlsx")


# open all the files in a folder to copy from
# pasteFile= input("File to paste into:")

# copyFolder= input("Folder to be copied: ")
# print(copyFolder)
# os.chdir(copyFolder)
# f= os.listdir(copyFolder)
# for i in f:
#     if i.endswith('.xlsx'):
#         createData(i)
