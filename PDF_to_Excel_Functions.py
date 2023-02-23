import pdfplumber #Used for converting the pdf into a text file
import os #Used to operate with external files such as extracting the PDF file names and deleting the text documents.

"""
THE FOLLOWING FUNCTIONS ARE FOR CONVERTING THE PDF INTO A TEXT FILE
START PDF TO TEXT FUNCTIONS
"""
def ConvertPDFtoText(pathPDF, inputPDF, path): #Parameters are the path of the PDF forms (pathPDF), the path of the current PDF (inputPDF), and the base path (path)
    #STEP 1: Create text file and open it in writing mode.
    textFile = path + "\\Touch_Point_re_Scraping_PDFs\\Text Files\\" + inputPDF + ".txt" #This is the name of the file plus where it's stored
    file = open(textFile, 'w')

    #STEP 2: Extract text from PDF
    pdf = pdfplumber.open(pathPDF)
    for c in range(2): #The 3rd page is not be included since that information is inconsistent
        page = pdf.pages[c] #page is assigned to the current page of the PDF
        text = page.extract_text() #The text is extracted from page
        file.write(text) #The text is written to the file

    #STEP 3: Properly format the text file.
    file = open(textFile, 'r')
    fileLines = file.readlines() #The text file is open in read mode first to set it's contents to a variable
    file = open(textFile , 'w')                        
    for line in fileLines: #Afterwards, we loop through all of those lines                       
        if isCertainLines(line) == False: #We pass the current line to isCertainLines to get rid of 'bad' lines. Including spaces.
            file.write(" ".join(line.split()) + "\n") #If the line is good, we re-write the line into the text file.

    #STEP 4: Close all opened files and return the text file
    file.close
    pdf.close
    return textFile

#This function is meant to cut out specific lines that aren't needed for populating the excel sheet. This significantly
#helps with the excel populating process.
def isCertainLines(line): #Parameter is the current line in the text document (line)
    if line.lstrip().rstrip().replace(" ", "").lower() == "customdatarequestform/formulairededemandededonnéespersonnalisées" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "datarequestdetails/détailssurlademandededonnées" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "specifyadaterangeforthedateextract/précisezuneplagededatespourl'extractiondesdonnées" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "requestordetails/détailssurledemandeur" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "4.billingaddress/adressedefacturation" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "billingaddress/adressedefacturation" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "4.billinginformation/renseignementssurlafacturation" or \
    line.lstrip().rstrip().replace(" ", "").lower() == "5.yourdetails/vosrenseignements" or \
    line.isspace():
        return True
    return False
"""
END PDF TO TEXT FUNCTIONS
"""

"""
THE FOLLOWING FUNCTIONS ARE FOR WRITING THE TEXT IN THE TEXT FILE INTO THE EXCEL DOCUMENT.
START OF TEXT TO EXCEL FUNCTIONS
"""
def ConvertTexttoExcel(textFile, sheet, row, portalNumber, date): #Parameters are the text file being written into excel (textFile), the main excel document (sheet), the currnet row being populated (row), the portal number of the form (portalNumber), and the date of the form (date)
    inputFile = open(textFile, 'r')
    fileLines = inputFile.readlines() #fileLines are the lines being read from the text file
    headerList = Headers(sheet) #headersList is the array containing the excel headers.

    #Seperate lines for populating portal number and date since they are not retrievable from the text file
    sheet.cell(row=row, column=1).value = portalNumber
    sheet.cell(row=row, column=2).value = date

    #This loop goes through the text file and begis populating the excel sheet
    l = 0
    while l < len(fileLines):
        try:
            c = int(HeaderValue(fileLines[l], headerList)) + 3
            x = 1
            while (HeaderValue(fileLines[l + x], headerList) == None):
                if sheet.cell(row=row, column=c).value == None: #If cell is blank, we set it as the line
                    sheet.cell(row=row, column=c).value = str(fileLines[l + x])
                else: #Otherwise, we can concatenate. We cannot concatenate with a 'None' value.
                    sheet.cell(row=row, column=c).value = sheet.cell(row=row, column=c).value + str(fileLines[l + x])
                x += 1
                if (l + x) >= len(fileLines):
                    break
            headerList[c - 3] = None
            l += x
        except:
            break
    return


#Function to loop through the first 50 cells to see if they are all blank
def CurrentBlankRow(sheet, row, min, max): #Parameters are the excel file (sheet), the first populatable row of the excel document (row), the first tested column (min), and the last tested column (max)
    r = row
    while True:
        c = min
        while c <= max: #Last c value is 50 because those columns are the ones being populated
            if sheet.cell(row=r, column=c).value != None: #If a cell is not empty, we break out the loop to go to the next row
                break
            c += 1
        if c == (max + 1): #If c is 51, it means that cells 1-50 are empty; which is the requirement.
            return r
        else: #If c isn't 51, the nested loop broke when it reached a cell with content in it. So r increments by 1 for the next row.
            r += 1

#This function takes the date and returns it properly formatted. Even if the date isn't properly formatted, the first 9
#characters are all we need. Strings after those 9 characters are ignored
def FormatDate(date): #Parameters are the portal date (date)
    month = str(date[0]) + str(date[1]) + str(date[2]) #First 3 character should be the month
    day = str(date[3]) + str(date[4]) #The next 2 characters should be the day
    year = str(date[5]) + str(date[6]) + str(date[7]) + str(date[8]) #The last 4 characters should be the year
    return month + " " + day + " " + year #The full date is returned with spaces in between

#Headers(sheet) returns an array containing all of the headers from the excel sheet. Does not include the custom data at the right side.
def Headers(sheet): #Parameters are the excel file (sheet)
    columnHeaders = [] #columnHeaders will be the array we return.
    c = 3 #We start on column 3 because we are manually populating the portal number and date
    currentCell = str(sheet.cell(row=3, column=c).value).lstrip().rstrip() #currentCell is the current cell being added to the array.
    while c <= 50:
        columnHeaders.append(currentCell)
        c += 1
        currentCell = str(sheet.cell(row=3, column=c).value).lstrip().rstrip()
    return columnHeaders

def HeaderValue(line, headerList): #Parameters are the current line in the text file (line), and the list of headers in the excel document (headerList)
    for i in range(len(headerList)):
        if str(headerList[i]).replace(" ", "").lower() in str(line).replace("\n","").replace(" ", "").lower().lstrip().rstrip() \
            and str(headerList[i]) != None:
            return i
    return None
"""
END OF TEXT TO EXCEL FUNCTIONS
"""

"""
THESE FUNCTIONS ARE FOR RETRIEVING THE PDFS AND DELETING THE TEXT DOCUMENTS THEY REQUIRE USE OF THE OPERATING SYSTEM
START OS FUNCTIONS
"""
#This function gets the names of every PDF and puts them into an array.
def ExtractAllPDFs(pdfPath): #Parameters are the path of the PDF forms (pdfPath)
    PDFList = []
    directory = os.fsencode(pdfPath)
    for file in os.listdir(directory):
        currentFile = os.fsencode(file)
        if ".pdf" in str(currentFile):
            PDFList.append(str(currentFile.decode('UTF-8'))) #Converts the PDF name into a normal string
    return PDFList

#This function deletes all of the temporary text files that are being stored.
def DeleteTextFiles(pdfPath): #Parameters are the path of the PDF forms (pdfPath).
    textFilePath = pdfPath + "\\Text Files" #The path of the text files adds on to the path of the PDF forms.
    directory = os.fsencode(textFilePath)
    for file in os.listdir(directory):
        currentFile = os.fsencode(file)
        if ".txt" in str(currentFile):
            os.remove(textFilePath + "\\" + str(currentFile.decode('UTF-8'))) #Removes the current text file
    return

"""
END OS FUNCTIONS
"""

"""
THESE FUNCTIONS ARE FOR FINDING THE DIFFERENCE OF DAYS BETWEEN COLUMNS B AND N
"""
def DifferenceDays(sheet, r, diff, d1, d2):
    dateB = str(sheet.cell(row=r, column=d1).value)
    dateN = str(sheet.cell(row=r, column=d2).value)
    if dateB != "None" and dateN != "None":
        dateBArray = dateB.split()
        dateNArray = dateN.split()
        dateBInt = int(ConvertMonthtoDays(str(dateBArray[0]), int(dateBArray[2]))) + int(dateBArray[1])
        dateNInt = int(ConvertMonthtoDays(str(dateNArray[0]), int(dateNArray[2]))) + int(dateNArray[1])
        dateDifference = (dateNInt - dateBInt) + ConvertYearstoDays(dateBArray[2], dateNArray[2])
        sheet.cell(row=r, column=diff).value = abs(int(dateDifference))
    else:
        sheet.cell(row=r, column=diff).value = "N/A"
    return

def ConvertMonthtoDays(month, year): #Parameters are the month (month) and year (year) of the pdf file date
    n = 0
    if month.lower() == "jan":
        pass
    else:
        n += 31
        if month.lower() == "feb":
            pass
        else:
            if year % 4 == 0:
                n += 29
            else:
                n += 28
            if month.lower() == "mar":
                pass
            else:
                n += 31
                if month.lower() == "apr":
                    pass
                else:
                    n += 30
                    if month.lower() == "may":
                        pass
                    else:
                        n += 31
                        if month.lower() == "jun":
                            pass
                        else:
                            n += 30
                            if month.lower() == "jul":
                                pass
                            else:
                                n += 31
                                if month.lower() == "aug":
                                    pass
                                else:
                                    n += 31
                                    if month.lower() == "sep":
                                        pass
                                    else:
                                        n += 30
                                        if month.lower() == "oct":
                                            pass
                                        else:
                                            n += 31
                                            if month.lower() == "nov":
                                                pass
                                            else:
                                                n += 30
                                                if month.lower() == "dec":
                                                    pass
                                                else:
                                                    return "Bad Input"
    return n

def ConvertYearstoDays(dateBYear, dateNYear):
    x = 0
    n = int(dateBYear)
    m = int(dateNYear)
    if n <= m:
        while n < m:
            if n % 4 == 0:
                x += 366
            else:
                x += 365
            n += 1
    else:
        while m < n:
            if m % 4 == 0:
                x -= 366
            else:
                x -= 365
            m += 1
    return x

"""
END DAYS DIFFERENCE FUNCTIONS
"""

"""
START FUNCTIONS COMPARING CELLS
"""
def CompareCells(sheet, sheetTwo, r1, r2):
    cellMain = str(sheet.cell(row=r1, column=1).value)
    cellForm = str(sheetTwo.cell(row=r2, column=3).value)
    if cellMain == cellForm:
        PopulateFormResults(sheet, sheetTwo, r1, r2)
    return

def PopulateFormResults(sheet, sheetTwo, r1, r2):
    c1 = 57
    c2 = 3
    while c1 < 70:
        sheet.cell(row=r1, column=c1).value = str(sheetTwo.cell(row=r2, column=c2).value)
        c1 += 1
        c2 += 1
    return
"""
END FUNCTION COMPARING CELLS
"""

"""
START FORMAT MONTH YEAR FUNCTIONS
"""
def FormatMonthYear(sheet, r):
    dateArray = sheet.cell(row=r, column=2).value.split()
    sheet.cell(row=r, column=72).value = str(FormatMonthText(dateArray[0])) + "/" + str(dateArray[2])
    return

def FormatMonthText(month):
    if str(month).lower() == "jan":
        return "01"
    elif str(month).lower() == "feb":
        return "02"
    elif str(month).lower() == "mar":
        return "03"
    elif str(month).lower() == "apr":
        return "04"
    elif str(month).lower() == "may":
        return "05"
    elif str(month).lower() == "jun":
        return "06"
    elif str(month).lower() == "jul":
        return "07"
    elif str(month).lower() == "aug":
        return "08"
    elif str(month).lower() == "sep":
        return "09"
    elif str(month).lower() == "oct":
        return "10"
    elif str(month).lower() == "nov":
        return "11"
    elif str(month).lower() == "dec":
        return "12"
    return "--"

def CountCombinations(sheet, r, currentRow, inputColumn, requestorType):
    comboDict = {}
    n = 0
    while r < currentRow:
        currentKey = str(sheet.cell(row=r, column=inputColumn).value)
        if requestorType + "\n" == str(sheet.cell(row=r, column=18).value) or requestorType == "All":
            if currentKey in comboDict:
                comboDict[currentKey] += 1
            else:
                comboDict[currentKey] = 1
            n += 1
            print("Row " + str(r) + " Counted")
        r += 1
    print("Number of elements: " + str(n))
    return comboDict
"""
END FORMAT MONTH YEAR FUNCTIONS
"""