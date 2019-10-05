from copybydictXlsx import createData,copyRange, pasteRange, searchForWord
import os
#Prepare the spreadsheets to copy from and paste.
# open all the files in a folder to copy from
# pasteFile= input("Paste File:")


# copyFolder= input("Copy Folder: ")
# os.chdir(copyFolder)
# f= os.listdir(copyFolder)
print(createData("01-09-2019.xlsx", "Sample2.xlsx"))

# for iterCopy in f:
    # if iterCopy.endswith('.xlsx'):
    #     createData(iterCopy, pasteFile)
    #     # print(iterCopy)