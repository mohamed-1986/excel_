from copybydictXlsx import moveData, TheSheets
# from copybydictXlsx import copyRange, pasteRange, searchForWord, TheSheets, searchRowStarting, dictionaries
import os
# import xlrd, openpyxl

#Prepare the spreadsheets to copy from and paste.
# open all the files in a folder to copy from

pasteFileName = input("Paste File:")
pasteFileSheet= "History 02"
copyFolder= input("Copy Folder: ")
os.chdir(copyFolder)

ff= os.listdir(copyFolder)
f=[sf for sf in ff if sf.endswith('.xls') or sf.endswith('xlsx')]

for copyFileName in f:
    wb, wss= TheSheets(copyFileName)
    for copyFileSheet in wss:
        print(copyFileName, copyFileSheet, pasteFileName ,"History 02")
        moveData(copyFileName, copyFileSheet, pasteFileName, pasteFileSheet)
    