import xlsxwriter 
import pypdf
import os


"""GLOBAL Variables"""
# Common words that are being filtered out of the document
filterEnteries = ["LSK", "CWB", "CWE", "CWA",
                  "Teacher(s)", "Room: 413", "Wednesday", 
                  "Number Student Name Phone Number", "Father Work Phone", "Mother Work Phone", "Guardian Work Phone",
                  "High School", "Students"]

pdfPages = []               # Each index is a page of the pdf
studentNumbers = []         # Student numbers that have been added to the excel file, due to multiple students taking multiple classes
duplicateStudent = False    # Global variable that is used if current student is already in the excel file
columnWidth = [2]*13        # Width that each column is set at
worksheet = None


curDir = os.getcwd()                            # Current directory
excelFolder = curDir + "\output"                # Directory where the excel files go
excelFileName = "pdf_data.xlsx"            # Name of the excel file
excelPath = excelFolder + "\\" + excelFileName  # Path to the excel file
#pdfName = "Lashyn.pdf"                          # Name of the pdf to analyze
pdfName = "PDF_to_XLSX_Mock_Data.pdf"
# Updates the stored list of column lengths
def updateWidthValue(content, column):

    index = ord(column) - 65

    # Update stored column length if content is larger
    if (len(content) > columnWidth[index]):
        columnWidth[index] = len(content)


# Sets the column length in the excel book, based on the global index
def setColumnWidth(worksheet):

    column = 'A'
    
    for width in columnWidth:
        worksheet.set_column(column + ':' + column, width + 2)
        column = chr(ord(column) + 1)


# Reads the pdf file and filters out the pointless information
def getPDFData(pdfPath):

    # Clears the PDF    
    global pdfPages
    pdfPages = [] 

    # creating a pdf reader object
    reader = pypdf.PdfReader(pdfPath)

    pageCount = 0

    # Adds each page to the pdf list
    for page in range (0, reader.get_num_pages()):
        pageContent = reader.pages[page].extract_text().split("\n")

        # Filter out the useless information from the page
        pageContent = [content for content in pageContent if not any(filter in content for filter in filterEnteries)]
        
        # Add the page content to the global list
        pdfPages.extend(pageContent)
        pageCount += 1
    print("Completed scanning " +str(pageCount) + " page(s)")

# Create the folder for the excel file if it doesn't already exist
def createFolder():
    if(not os.path.exists(excelFolder)):
        os.makedirs(excelFolder)
        print("Created new Folder")
    else:
        print("Folder already exist")


# Checks if the student number was already inputed into the system, this avoid duplicates
def checkStudentNum(info):

    splitInfo = info.split(" ")
    studentNumber = splitInfo[0]
    global duplicateStudent

    if(studentNumber not in studentNumbers):
        studentNumbers.append(studentNumber)
        duplicateStudent = False
        return
    
    duplicateStudent = True


# Adds the student info to the excel file
def addStudentInfo(info, worksheet, row):

    studentColumn = ['A', 'B', 'C', 'D', 'E']

    info = info.replace(") ", ")")
    info = info.replace(",", "")
    info = info.rstrip()
    splitInfo = info.split(" ")
    
    hasMiddle = len(splitInfo) == 5
    counter = 0
    for column in studentColumn:

        # Checks that there is enough data in the loop
        if(counter >= len(splitInfo)):
            break

        # Skips if the user doesn't have a middle name
        if (column == "D" and not hasMiddle):
            continue

        worksheet.write(column + str(row), splitInfo[counter])
        updateWidthValue(splitInfo[counter],column)

        counter += 1

    return
  

# Function to add the guardian info to the excel file
def addGuardianInfo(info, worksheet, row, guardian):
    
    if(guardian == 1):
        column = 'F'
    elif (guardian == 2):
        column = 'J'
    else:
        return

    # Splits the guardian info has a number attached
    guardianInfo = info.split(": (")

    worksheet.write(column + str(row), guardianInfo[0])
    updateWidthValue(guardianInfo[0],column)
    # Attempts to add the guardian phone number
    try:
        column = chr(ord(column) + 1)
        cell = "("+ guardianInfo[1]
        cell = cell.replace(") ", ")")
        cell = cell.replace(")", ") ")
        worksheet.write(column + str(row), cell)
        updateWidthValue(cell,column)
    except:
        return


# Function to add the guardian cell number
def addGuardianCell(info, worksheet, row, guardian):

    if(guardian == 1):
        column = "H"
    elif (guardian == 2):
        column = "L"
    else:
        return
    
    info = info.replace("Cell:", "")
    info = info.replace(") ", ")")
    info = info.replace(")", ") ")
    info = info.lstrip()
    worksheet.write(column + str(row), info)
    updateWidthValue(info,column)

# Function to add the guardian email address
def addGuardianEmail(info, worksheet, row, guardian):
    
    if(guardian == 1):
        column = "I"
    elif (guardian == 2):
        column = "M"
    else:
        return

    info = info.replace("Email:", "")
    info = info.lstrip()
    worksheet.write(column + str(row), info)
    updateWidthValue(info,column)


# Creates new excel file
def createExcelFile():

    workbook = xlsxwriter.Workbook(excelPath)
    global worksheet
    worksheet = workbook.add_worksheet() 
    headerInfo = ['Number', 'Student Last Name', 'Student First Name', 'Student Middle Name','Phone Number',
                  'Guardian 1', 'Home Number', 'Work Phone Number', 'Email',
                  'Guardian 2', 'Home Number', 'Work Phone Number', 'Email',]
    column = 'A'
    for header in headerInfo:
        worksheet.write(column + '1', header)
        updateWidthValue(header,column)
        column = chr(ord(column) + 1)

    print("Excel file created")
    return workbook

# Loops through the stored pdf info and adds it excel file
def addInfo():
    row = 1
    guardian = 0
    for info in pdfPages:
        
        ## Checks 
        if(info[0] == "0"):
            checkStudentNum(info)

            if(duplicateStudent):
                continue
            row += 1
            guardian = 0
            addStudentInfo(info, worksheet, row)
            
            

        # Guardian(s) phone number
        elif(info.find("Cell:") != -1):
            if(duplicateStudent):
                continue
            addGuardianCell(info, worksheet, row, guardian)
            
        elif(info.find("Email:") != -1):
            if(duplicateStudent):
                continue
            addGuardianEmail(info, worksheet, row, guardian)

        else:
            guardian += 1
            addGuardianInfo(info, worksheet, row, guardian)     

    

if __name__ == "__main__":
    pdfName = input("Enter the PDF Name: ").strip()
    createFolder()
    excelFile = createExcelFile()
    getPDFData(curDir + "\\"+ pdfName)
    addInfo()
    setColumnWidth(worksheet)

    try:
        excelFile.close()
        print("Data saved to " + str(excelFileName))
    except:
        print("Error: Unable to save data to excel file")
