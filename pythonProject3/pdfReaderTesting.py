# importing required modules
import PyPDF2
import os
import xlsxwriter
from tkinter import*
import tkinter as tk

py = []
master = Tk()
master.title("CilasPal: Grain Size Digitizer")
master.geometry("450x250")

e = Entry(master)
e.pack()
e.focus_set()
var = StringVar()
label = Label(master, textvariable=var, relief=RAISED, wraplength=400)

var.set("Please paste the directory from which you would like to retrieve data by doing the following 1. Navigate to the folder of the cilas files on this computer (this will most likely be located on an external drive) 2. In the file destination field (the space to the left of the file search field), left click once to highlight the directory 3. Copy and paste here and click OK")
label.pack()


def callback():
    repo = e.get() # This is the text you may want to use later
    print("retrieving data from repository...", repo)
    py.append(repo)
    master.destroy()

b = Button(master, text = "OK", width = 10, command = callback)
b.pack()
mainloop()

#  the initialization components of the spreadsheet to sort inserted data
insList = ["ID", "Mean", "Median", 0.04, 0.07, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1,
           1.2, 1.3, 1.4, 1.6, 1.8, 2.0, 2.2, 2.4, 2.6, 3.0, 4.0, 5.0, 6.0, 6.5, 7.0,
           7.5, 8.0, 8.5, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0,
           19.0, 20.0, 22.0, 25.0, 28.0, 32.0, 36.0, 38.0, 40.0, 45.0, 50.0, 53.0, 56.0,
           63.0, 71.0, 75.0, 80.0, 85.0, 90.0, 95.0, 100.0, 106.0, 112.0, 125.0, 130.0,
           140.0, 145.0, 150.0, 160.0, 170.0, 180.0, 190.0, 200.0, 212.0, 242.0, 250.0,
           300.0, 400.0, 500.0, 600.0, 700.0, 800.0, 900.0, 1000.0, 1100.0, 1200.0, 1300.0,
           1400.0, 1500.0, 1600.0, 1700.0, 1800.0, 1900.0, 2000.0, 2100.0, 2200.0, 2300.0,
           2400.0, 2500.0]
#creating a list for UDSC
insList2 = [0.04, 3.90, 62.00, 88.00, 125.00, 177.0, 2500.0, 350.0, 500.0, 710.0, 1000.0, 1410.0, 2000.0]

# Creates a workbook in excel to house all the data
workbook = xlsxwriter.Workbook(py[0] + '//YourData.xlsx')
worksheet = workbook.add_worksheet("Primary Data Tables")
worksheet2 = workbook.add_worksheet("User Defined Size Classes")

# worksheet.write (row from 0, col from 0, item) ... pastes initialization components
for i in range(103):
    worksheet.write(0, i, insList[i])
for i in range(13):
    worksheet2.write(0, i + 1, insList2[i])

# why wont this work repo = input("Paste Your Repository Here: ") --> becasue using os.listdir and not open()?
# repo = input("Please input the repository of your cilas data here: ")

print("\033[1;31mRepository received  \n" )
print("Error:GetFileNotFound IGNORED")
pdfsList = []

# 1.Get file names from directory
file_list = os.listdir(py[0])

# 2. generates a list of all the pdfs in the repository
for i in range(len(file_list)):
    file = file_list[i]
    if file.__contains__('.pdf'):
        pdfsList.append(file)
print(pdfsList)

for p in range(len(pdfsList)):
    #   creating a pdf file object
    pdfFileObj = open(py[0] + '/' + pdfsList[p], 'rb')

    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict="false")

    # printing number of pages in pdf file
    # print(pdfReader.numPages)

    # creating a page object
    pageObj = pdfReader.getPage(1)
    pageObj2 = pdfReader.getPage(0)

    List = pageObj.extractText()
    List2 = pageObj2.extractText()

    littleQArray = []
    bigQArray = []
    Index = List2.index("undersizexQ3")
    startIndex = Index + 19
    endIndex = startIndex + 6

    # Extracting the User defined size classes
    for i in range(4):
        bigQArray.append(List2[startIndex:endIndex])
        startIndex = startIndex + 13
        endIndex = endIndex + 13

    startIndex = startIndex - 1
    endIndex = endIndex - 1  
    for i in range(6):
        bigQArray.append(List2[startIndex:endIndex])
        startIndex = startIndex + 12
        endIndex = endIndex + 12

    startIndex = startIndex + 5
    endIndex = endIndex + 5
    for i in range(3):
        bigQArray.append(List2[startIndex:endIndex])
        startIndex = startIndex + 13
        endIndex = endIndex + 13


    # Extracting the Full Size Class
    sIndex = List.index('undersize')+ 28
    eIndex = sIndex + 6
    for i in range(6):
        for j in range(10):
            ins = List[sIndex:eIndex]
            littleQArray.append(ins)
            sIndex = sIndex + 19
            eIndex = eIndex + 19

        sIndex = sIndex + 6
        eIndex = eIndex + 6

    sIndex = sIndex + 0
    eIndex = eIndex + 0
    for i in range(2):
        for j in range(10):
            ins = List[sIndex:eIndex]
            littleQArray.append(ins)
            sIndex = sIndex + 18
            eIndex = eIndex + 18
        sIndex = sIndex + 6
        eIndex = eIndex + 6

    for j in range(4):
        ins = List[sIndex:eIndex]
        littleQArray.append(ins)
        sIndex = sIndex + 18
        eIndex = eIndex + 18

    sIndex = sIndex + 1
    eIndex = eIndex + 1
    for j in range(6):
        ins = List[sIndex:eIndex]
        littleQArray.append(ins)
        sIndex = sIndex + 19
        eIndex = eIndex + 19
    sIndex = sIndex + 6
    eIndex = eIndex + 6
    for j in range(10):
        ins = List[sIndex:eIndex]
        littleQArray.append(ins)
        sIndex = sIndex + 19
        eIndex = eIndex + 19

    # retrieving mean and median
    Apples = List.index('Mean diameter : ')
    ApplesFinal = Apples + 16
    mean = List[ApplesFinal: ApplesFinal + 6]
    median = List[Apples - 37 :Apples - 30]

    # retrieving sample name
    nameIndexStart = List.index('Sample ref.')
    nameIndexFinish = List.index('Sample Name')
    sampleName = List[nameIndexStart + 13: nameIndexFinish]

    littleQArray.insert(0, sampleName)
    bigQArray.insert(0, sampleName)
    littleQArray.insert(1, mean)
    littleQArray.insert(2, median)

    # printing info to console for error detection
    print(sampleName, littleQArray, mean, median, bigQArray, sep = ",")

    for a in range(len(littleQArray)):
        worksheet.write(p + 1, a, littleQArray[a])

    for a in range(len(bigQArray)):
        worksheet2.write(p + 1, a, bigQArray[a])

    # closing the pdf file object
    pdfFileObj.close()
workbook.close()

# Notify the user when the program finishes running
final = Tk()
final.title("CilasPal: Grain Size Digitizer")
final.geometry("450x250")


var = StringVar()
label = Label(final, textvariable=var, relief=RAISED, wraplength=400)

var.set("Your Excel file (YourData.xlsx) has been created. You may view your data in the previously input file loaction. Close this window to close the app")
label.pack()
mainloop()
