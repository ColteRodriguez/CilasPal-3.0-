import PyPDF2

pdfFileObj = open(r'C:\Users\Tony\Dropbox\PythonScreenShotTest/BS_3-2.pdf', 'rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

pageObj = pdfReader.getPage(0)

List = pageObj.extractText() # extracting text from page 0

bigQArray = []

Index = List.index("undersizexQ3")
startIndex = Index + 19
endIndex = startIndex + 6

for i in range(4):
    bigQArray.append(List[startIndex: endIndex])
    startIndex = startIndex + 13
    endIndex = endIndex + 13

startIndex = startIndex - 1
endIndex = endIndex - 1
for i in range(6):
    bigQArray.append(List[startIndex:endIndex])
    startIndex = startIndex + 12
    endIndex = endIndex + 12

startIndex = startIndex + 5
endIndex = endIndex + 5
for i in range(3):
    bigQArray.append(List[startIndex: endIndex])
    startIndex = startIndex + 13
    endIndex = endIndex + 13

print(bigQArray)


print(List)

pdfFileObj.close()