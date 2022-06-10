import base64
import os
import re
from zipfile import ZipFile
import pandas as pd
import pdfplumber


def adandeFileDownload(fileA, filePath):
    fileName = fileA['name']
    # Checks to see if file is a zip
    if re.search(r'\.zip$', fileName) != None or re.search(r'\.ZIP$', fileName) != None:
        # if it's a zip, save it
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        fileNameShort = fileName.replace('.zip', "ZIP")
        # unzip the folder
        temp = os.path.join(filePath, fileName)
        with ZipFile(temp, 'r') as zipObj:
            zipObj.extractall(os.path.join(filePath + "\\" + fileNameShort))
        # loop through files
        directory = os.path.join(filePath, fileNameShort)
        for filename in os.listdir(directory):
            filenameSplit = filename.split(".")
            filenameShort = re.search(r"(\d*)-(\d*)", filenameSplit[0])
            filenameExtension = filenameSplit[1]
            filenameSplit = filenameShort.group(0).split("-")
            rev = filenameSplit[1]

            filenameNew = filenameSplit[0] + " Rev " + str(rev) + "." + filenameExtension
            f = os.path.join(directory, filename)
            # checking if it is a file
            if os.path.isfile(f):
                # remove the extension to creat the folder name
                folderName = filenameSplit[0]
                # Save the file
                newpathFolder = filePath + "\\" + str(folderName)
                # Create the file path if it doesn't already exist
                if not os.path.exists(newpathFolder):
                    os.makedirs(newpathFolder)
                # Cut the file to the desired locations
                if not os.path.exists(newpathFolder + "\\" + filenameNew):
                    os.rename(os.path.join(directory, filename), newpathFolder + "\\" + filenameNew)
                else:
                    os.remove(os.path.join(directory, filename))
        # removed the save files that are no longer needed
        os.remove(os.path.join(filePath, fileName))
        os.rmdir(os.path.join(filePath, fileNameShort))
        print("All the Files have been saved for: " + str(fileName))


def bakerPerkinsFileDownload(fileA, filePath):
    fileName = fileA['name']
    moreThenOneSheet = False
    # Checks to see if file is a pdf or dxf
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        # check to see if it should be more then one sheet
        sheetNumber = re.search(r"SH(\d*)", fileName)
        if sheetNumber != None:
            # if more then one sheet set moreThenOnSheet true and extract the sheet number
            moreThenOneSheet = True
            sheetNumber = sheetNumber.group(0)
            sheetNumber = sheetNumber[2:]
            if sheetNumber.isdigit():
                sheetNumber = int(sheetNumber)
        # Check to see if it contains an underscore
        fileNameSplit = ""
        partNumber = ""
        revision = "Unknown"
        # Locate the partnumber and revision
        if re.search(r"_", fileName) != None:
            # if it does set partNumber and revision
            fileNameSplit = fileName.split("_")
            partNumber = fileNameSplit[0]
            revision = fileNameSplit[1]
        else:
            fileNameSplit = fileName.split(".")
            fileNameSplit = fileNameSplit[0].split(" ")
            partNumber = fileNameSplit[0]
            # else set revision to unknown
        # Create a new ame for the file including it's rev
        filenameSplit = fileName.split(".")

        # Assigning filenameNew, value depending on moreThenOneSheet
        if moreThenOneSheet == False:
            filenameNew = partNumber + " Rev " + str(revision) + "." + filenameSplit[1]
        else:
            filenameNew = str(partNumber) + " Rev " + str(revision) + " Sheet " + str(sheetNumber) + "." + \
                          filenameSplit[1]

        # Save the file
        newpathFolder = filePath + "\\" + str(partNumber)
        # Create a folder for the part if that folder does not exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the folder
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        print("File:" + filenameNew + " has been saved")


def bradmanLakeFileDownload(fileA, filePath, dfRev):
    if dfRev.empty:
        dfRev.append({'Part_Number': "Unknown", 'Revision': "Unknown"}, ignore_index=True)
    fileName = fileA['name']
    # Checks to see if file is a zip
    if re.search(r'\.zip$', fileName) != None or re.search(r'\.ZIP$', fileName) != None:
        # if it's a zip, save it
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        fileNameShort = fileName.replace('.zip', "")
        # unzip the folder
        temp = os.path.join(filePath, fileName)
        with ZipFile(temp, 'r') as zipObj:
            zipObj.extractall(os.path.join(filePath + "\\" + fileNameShort))
        # loop through files

        directory = os.path.join(filePath, fileNameShort)
        for filename in os.listdir(directory):
            filenameSplit = filename.split(".")

            rev = "unknown"
            for index, row in dfRev.iterrows():
                if row['Part_Number'] == filenameSplit[0]:
                    rev = row['Revision']

            filenameNew = filenameSplit[0] + " Rev " + str(rev) + "." + filenameSplit[1]
            f = os.path.join(directory, filename)
            # checking if it is a file
            if os.path.isfile(f):
                # remove the extension to creat the folder name
                folderName = f.replace(directory + "\\", "").split(".")
                folderName = folderName[0]
                folderName = folderName.replace(directory + "\\", "").split("_")
                folderName = folderName[0]
                # Save the file
                # for ind in df_Customer_Regex.index:
                # if df_Customer_Regex['Customer'][ind] == "BRADMAN-LAKE":
                newpathFolder = filePath + "\\" + str(folderName)
                # Create the file path if it doesn't already exist
                if not os.path.exists(newpathFolder):
                    os.makedirs(newpathFolder)
                # Cut the file to the desired locations
                if not os.path.exists(newpathFolder + "\\" + filenameNew):
                    os.rename(os.path.join(directory, filename), newpathFolder + "\\" + filenameNew)
                else:
                    os.remove(os.path.join(directory, filename))
        # removed the save files that are no longer needed
        os.remove(os.path.join(filePath, fileName))
        os.rmdir(os.path.join(filePath, fileNameShort))
        print("All the Files have been saved for: " + str(fileName))
    elif (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        fileName = fileA['name']
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        fileNameShort = fileName.replace('.dxf', "")
        fileNameShort = fileNameShort.replace('.pdf', "")
        folderName = fileNameShort.split("_")
        folderName = folderName[0]

        # Create a new ame for the file including it's rev
        filenameSplit = fileName.split(".")

        rev = "unknown"
        for index, row in dfRev.iterrows():
            if row['Part_Number'] == filenameSplit[0]:
                rev = row['Revision']

        filenameNew = filenameSplit[0] + " Rev " + str(rev) + "." + filenameSplit[1]

        # Save the file
        newpathFolder = filePath + "\\" + str(folderName)
        # Create the file path if it doesn't already exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the desired locations
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        # removed the save files that are no longer needed
        # os.remove(os.path.join(filePath, fileName))
        # os.rmdir(os.path.join(filePath, fileNameShort))
        print("File:" + filenameNew + " has been saved")


def bradmanLakeRevTableCreator(pdfSavedFiled):
    df = pd.DataFrame(columns=['Part_Number', 'Revision'])
    lineCounter = 1
    pdf = pdfplumber.open(pdfSavedFiled)
    for page in pdf.pages:
        for line in page.extract_text().split("\n"):
            partNumber = ""
            partRev = ""
            splitLine = line.split(" ")
            if splitLine[0].isdigit():
                if int(splitLine[0]) == lineCounter:
                    pattern = r"^SC\-\d*"
                    if (re.search(pattern, splitLine[1]) is None):
                        pass
                    else:
                        splitLine[1] = splitLine[1].replace("SC-", "")
                    partNumber = splitLine[1]
                    pattern = r"^(\d*)$|^F(\d*)|(\d*)-(\d*)\w$"
                    if (re.search(pattern, splitLine[1]) is None):
                        pass
                    else:
                        partRev = splitLine[2]
                        lineCounter = lineCounter + 1
                        df = df.append({'Part_Number': partNumber, 'Revision': partRev}, ignore_index=True)
    return df


def fordsFileDownload(fileA, filePath, drawingList):
    fileName = fileA['name']
    # Checks to see if file is a pdf or dxf or swg
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        for name in drawingList:
            if re.search(name, fileName) != None:
                fileNameSplit = ""
                partNumber = name
                revision = "Unknown"
                # Locate the partnumber and revision
                if re.search(r"-", fileName) != None:
                    # if it does set partNumber and revision
                    fileNameSplit = fileName.split("-")
                    partNumber = fileNameSplit[0]
                    revision = fileNameSplit[1]
                    # else set revision to unknown
                # Create a new name for the file including it's rev
                filenameSplit = fileName.split(".")
                # Assigning filenameNew, value depending on moreThenOneSheet
                filenameNew = str(partNumber) + " Rev " + str(revision) + "." + filenameSplit[1]

                # Save the file
                newpathFolder = filePath + "\\" + str(partNumber)
                # Create a folder for the part if that folder does not exist
                if not os.path.exists(newpathFolder):
                    os.makedirs(newpathFolder)
                # Cut the file to the folder
                if not os.path.exists(newpathFolder + "\\" + filenameNew):
                    os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
                else:
                    os.remove(os.path.join(filePath, fileName))
                print("File:" + filenameNew + " has been saved")


def fordsPartNumbersReturn(attachments):
    returnList = []
    for attachment in attachments:
        attachmentName = attachment['name']
        if re.search("DRAWING", attachmentName, ) != None:
            attachmentNameSplit = attachmentName.split("-")
            returnList.append(attachmentNameSplit[0])
    return returnList


def harrodFileDownload(fileA, filePath):
    fileName = fileA['name']
    # Checks to see if file is a pdf or dxf or swg
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        revision = "Unknown"
        filenameSplit = fileName.split(".")
        partNumber = filenameSplit[0]
        # Assigning filenameNew, value depending on moreThenOneSheet
        filenameNew = str(partNumber) + " Rev " + str(revision) + "." + filenameSplit[1]

        # Save the file
        newpathFolder = filePath + "\\" + str(partNumber)
        # Create a folder for the part if that folder does not exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the folder
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        print("File:" + filenameNew + " has been saved")


def hydraFileDownload(fileA, filePath):
    fileName = fileA['name']
    # Checks to see if file is a pdf or dxf or swg
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        revision = "Unknown"
        filenameSplit = fileName.split(".")
        partNumber = filenameSplit[0]
        # Assigning filenameNew, value depending on moreThenOneSheet
        filenameNew = str(partNumber) + " Rev " + str(revision) + "." + filenameSplit[1]

        # Save the file
        newpathFolder = filePath + "\\" + str(partNumber)
        # Create a folder for the part if that folder does not exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the folder
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        print("File:" + filenameNew + " has been saved")


def pharosFileDownload(fileA, filePath):
    fileName = fileA['name']
    # Checks to see if file is a pdf or dxf or dwg
    if (re.search(r"(\S*)-(\S*)-(\S*)\.(\w*)", fileName) == None):
        exit
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        revision = "Unknown"
        partNumber = ""
        if (re.search(r'.*(R\d)', fileName) != None):
            revision = re.search(r'.*(R\d)', fileName).group(1)
            partNumberSplit = fileName.split(" ")
            partNumber = partNumberSplit[0]
        else:
            filenameSplit = fileName.split(".")
            partNumber = filenameSplit[0]

        filenameSplit = fileName.split(".")
        # Assigning filenameNew, value depending on moreThenOneSheet
        filenameNew = str(partNumber) + " Rev " + str(revision) + "." + filenameSplit[1]

        # Save the file
        newpathFolder = filePath + "\\" + str(partNumber)
        # Create a folder for the part if that folder does not exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the folder
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        print("File:" + filenameNew + " has been saved")


def timberwolfFileDownload(fileA, filePath):
    fileName = fileA['name']
    filenameNew = ""
    # Checks to see if file is a pdf or dxf or dwg
    if (re.search(r"iss", fileName) == None) or (re.search(r"Iss", fileName) == None):
        exit
    if (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        ofValue = re.search(r'(\d*)\sof\s(\d*)', fileName).group(1)

        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        filenameSplit = fileName.split(" ")
        partNumber = filenameSplit[0]
        revision = filenameSplit[2]

        filenameSplit = fileName.split(".")
        if ofValue == None:
            filenameNew = str(partNumber) + " Rev " + str(revision) + "." + filenameSplit[1]
        else:
            filenameNew = str(partNumber) + " Rev " + str(revision) + " " + ofValue + "." + filenameSplit[1]
        # Save the file
        newpathFolder = filePath + "\\" + str(partNumber)
        # Create a folder for the part if that folder does not exist
        if not os.path.exists(newpathFolder):
            os.makedirs(newpathFolder)
        # Cut the file to the folder
        if not os.path.exists(newpathFolder + "\\" + filenameNew):
            os.rename(os.path.join(filePath, fileName), (newpathFolder + "\\" + filenameNew))
        else:
            os.remove(os.path.join(filePath, fileName))
        print("File:" + filenameNew + " has been saved")


def westrockFileDownload(fileA, filePath):
    fileName = fileA['name']
    quoteFileLocationCustomer = "C:\\\\Users\\josha\\PycharmProjects\\The Virtual Receptonist\\Test Storage\\"
    # Checks to see if file is a zip
    if re.search(r'^LS-(\s*)\.zip$', fileName) != None or re.search(r'^LS-(\s*)\.ZIP$', fileName):
        print("LS-zip")
        f = open(os.path.abspath(quoteFileLocationCustomer + fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        folderNameSplit = fileName.split(".")
        folderName = folderNameSplit[0]

        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        fileNameShort = fileName.replace('.zip', "ZIP")
        # unzip the folder
        temp = os.path.join(filePath, fileName)
        with ZipFile(temp, 'r') as zipObj:
            zipObj.extractall(os.path.join(filePath + "\\" + fileNameShort))
        # loop through files
        directory = os.path.join(filePath, fileNameShort)
        for filename in os.listdir(directory):
            filenameSplit = filename.split(".")
            filenameShort = filenameSplit[0]
            filenameExtension = filenameSplit[1]

            filenameNew = filenameSplit[0] + "." + filenameExtension
            f = os.path.join(directory, filename)
            # checking if it is a file
            if os.path.isfile(f):
                # Save the file
                newpathFolder = filePath + "\\" + str(folderName) + "\\"
                # Create the folder if it doesn't exist
                if not os.path.exists(newpathFolder):
                    os.makedirs(newpathFolder)
                # Cut the file to the desired locations
                if not os.path.exists(newpathFolder + "\\" + filenameNew):
                    filefrtg = newpathFolder + "\\" + filenameNew
                    print("filefrtg: " + filefrtg)
                    os.rename(os.path.join(directory, filename), newpathFolder + "\\" + filenameNew)
                else:
                    os.remove(os.path.join(directory, filename))
        # removed the save files that are no longer needed
        os.remove(os.path.join(filePath, fileName))
        os.rmdir(os.path.join(filePath, fileNameShort))
        print("All the Files have been saved for: " + str(fileName))
    elif re.search(r'\.zip$', fileName) != None or re.search(r'\.ZIP$', fileName) != None:
        print("zip")

        f = open(os.path.abspath(quoteFileLocationCustomer + fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()

        # if it's a zip, save it
        f = open(os.path.join(filePath, fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
        fileNameShort = fileName.replace('.zip', "ZIP")
        # unzip the folder
        temp = os.path.join(filePath, fileName)
        with ZipFile(temp, 'r') as zipObj:
            zipObj.extractall(os.path.join(filePath + "\\" + fileNameShort))
        # loop through files
        directory = os.path.join(filePath, fileNameShort)
        for filename in os.listdir(directory):
            filenameSplit = filename.split(".")
            filenameShort = filenameSplit[0]
            filenameExtension = filenameSplit[1]
            filenameNew = filename
            f = os.path.join(directory, filename)
            # checking if it is a file
            if os.path.isfile(f):
                # remove the extension to creat the folder name
                folderName = filenameSplit[0]
                # Save the file
                newpathFolder = filePath + "\\" + str(folderName)
                # Create the file path if it doesn't already exist
                if not os.path.exists(newpathFolder):
                    os.makedirs(newpathFolder)
                # Cut the file to the desired locations
                if not os.path.exists(newpathFolder + "\\" + filenameNew):
                    print(os.path.join(directory, filename))
                    print(newpathFolder + "\\" + filenameNew)
                    os.rename(os.path.join(directory, filename), newpathFolder + "\\" + filenameNew)
                else:
                    os.remove(os.path.join(directory, filename))
        # removed the save files that are no longer needed
        os.remove(os.path.join(filePath, fileName))
        os.rmdir(os.path.join(filePath, fileNameShort))
        print("All the Files have been saved for: " + str(fileName))

    elif (re.search(r'\.pdf$', fileName) != None) or (re.search(r'\.dxf$', fileName) != None) or (
            re.search(r'\.dwg$', fileName) != None) or (re.search(r'\.PDF$', fileName) != None) or (
            re.search(r'\.DXF$', fileName) != None) or (re.search(r'\.DWG$', fileName) != None):
        print("pdf")

        f = open(os.path.abspath(quoteFileLocationCustomer + fileName), 'w+b')
        f.write(base64.b64decode(fileA['contentBytes']))
        f.close()
