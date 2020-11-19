import PyPDF2
import os
from os import listdir, remove, path
from os.path import isfile, join
from os import walk
import hashlib
import sys
from datetime import datetime
import xml.etree.ElementTree as ET
import xlwt 
from xlwt import Workbook 

""" TODO: import this module, otherwise it's not safe
import defusedxml
import defusedxml.ElementTree as ET """

def getAttachments(reader):
    """
    Retrieves the file attachments of the PDF as a dictionary of file names
    and the file data as a bytestring.

    :return: dictionary of filenames and bytestrings
    """
    catalog = reader.trailer["/Root"]
    fileNames = catalog['/Names']['/EmbeddedFiles']['/Names']
    attachments = {}
    for f in fileNames:
        if isinstance(f, str):
            name = f
            dataIndex = fileNames.index(f) + 1
            fDict = fileNames[dataIndex].getObject()
            fData = fDict['/EF']['/F'].getData()
            attachments[name] = fData
    return attachments

def extractAttachments(path, filename, cleanup_after):
    if (filename[-3:].lower() == 'pdf'):
        handler = open(path + filename, 'rb')
        reader = PyPDF2.PdfFileReader(handler, strict=False)
        dictionary = getAttachments(reader)
        for fName, fData in dictionary.items():
            oFileName = path + filename[:-4] + '_' + fName
            with open(oFileName, 'wb') as outfile:
                outfile.write(fData)
            if (fName[-3:].lower() == 'xml'):
                aRecord = get_record(oFileName)
                if (aRecord):
                    md5_hash = hashlib.md5()
                    md5_hash.update(open(path + filename, "rb").read())
                    hhex = md5_hash.hexdigest()
                    aRecord['file_id'] = hhex
                    aRecord['file_source'] = path + filename
                    tbl.append(aRecord)
            if (cleanup_after):
                remove(oFileName)    
        return 1
    else:
        return 0

def get_record(iFileName):
    tree = ET.parse(iFileName)
    root = tree.getroot()
    if (root.tag == 'kezbesitesi_igazolas'):
        oRecord = {}
        for main_sec in root:
            for sub_sec in main_sec:
                #print(main_sec.tag + '_' + sub_sec.tag + ': ' + str(sub_sec.text))
                oRecord[main_sec.tag + '_' + sub_sec.tag] = str(sub_sec.text)
                if(main_sec.tag + '_' + sub_sec.tag in colnames.keys()):
                    colnames[main_sec.tag + '_' + sub_sec.tag] +=1 
                else:
                    colnames[main_sec.tag + '_' + sub_sec.tag] = 1
        return oRecord
    else:
        return None

def normalize_table(iTbl, iColnames):
    oTbl=[]
    for aRecord in iTbl:
        oRecord = {}
        for aCol in iColnames.keys():
            if (aCol in aRecord.keys()):
                oRecord[aCol] = aRecord[aCol]
            else:
                oRecord[aCol] = ''
        oTbl.append(oRecord)
    return oTbl

def table2csv(iTbl, iColnames, oFilename):
    aRow =''
    for aCol in iColnames:
        aRow = aRow + aCol +';'
    aRow = aRow[:-1] + '\n'
    with open(oFilename, 'w') as outfile:
        outfile.write(aRow)
        for aRecord in iTbl:
            aRow = ''
            for aCol in iColnames.keys():
                if (aCol in aRecord.keys()):
                    aRow = aRow + '"' + aRecord[aCol] + '"' + ';'
                else:
                    aRow = aRow + ';'
            aRow = aRow[:-1] + '\n'
            outfile.write(aRow)
    return True

def table2xls(iTbl, iColnames, oFilename):
    # Workbook is created 
    wb = Workbook() 
    # add_sheet is used to create sheet. 
    sheet1 = wb.add_sheet("kezb_"+ datetime.now().strftime("%m%d%y%H%M%S")) 
    row, col = 0, 0
    for aCol in iColnames:
        sheet1.write(row, col, aCol)
        col += 1
    with open(oFilename, 'w') as outfile:
        for aRecord in iTbl:
            col = 0
            row += 1
            for aCol in iColnames.keys():
                if (aCol in aRecord.keys()):
                    sheet1.write(row, col, aRecord[aCol])
                    col += 1
                else:
                    col += 1
    wb.save(oFilename) 
    return True

if (len(sys.argv) != 2):
    mypath, myfilename = os.path.split(os.path.abspath(__file__))
    mypath = mypath + '\\'
else:
    mypath = sys.argv[1]
    mypath = mypath.replace('\\\\', '\\')

onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
c=0
tbl=[]
colnames={}
for f in onlyfiles:
    c += extractAttachments(mypath, f, True)  # set the last param to False if you want the see the temporary xml files
print('Total of ' + str(c) +' .pdf files handled in the directory: ' )
colnames['file_id'] = c
colnames['file_source'] = c
tbl  = normalize_table(tbl,colnames)
#success = table2csv(tbl,colnames, mypath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".csv")
success = table2xls(tbl,colnames, mypath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".xls")

##TODO: data --> SQL
##TODO: get the cleanupafter prop as (command-line) arg
##TODO: if same-name column exists avoid collision (make a new colname "old2")
##TODO: define the default colnames by specification - ake the output universal even if there is no data in the xml for some columns
##TODO: excel export: https://www.geeksforgeeks.org/writing-excel-sheet-using-python/ VAGY https://xlsxwriter.readthedocs.io/ kerdes, hogy melyik more lightweight
##TODO: encrypted exe translation