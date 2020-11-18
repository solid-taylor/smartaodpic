import PyPDF2
from os import listdir
from os.path import isfile, join
from os import walk
import sys
import xml.etree.ElementTree as ET

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

def extractAttachments(path, filename):
    if (filename[-3:].lower() == 'pdf'):
        handler = open(path + filename, 'rb')
        reader = PyPDF2.PdfFileReader(handler)
        dictionary = getAttachments(reader)
        #print(dictionary)
        for fName, fData in dictionary.items():
            oFileName = path + filename[:-4] + '_' + fName
            with open(oFileName, 'wb') as outfile:
                outfile.write(fData)
            if (fName[-3:].lower() == 'xml'):
                handle_xml(oFileName)

def handle_xml(iFileName):
    tree = ET.parse(iFileName)
    root = tree.getroot()
    if (root.tag == 'kezbesitesi_igazolas'):
        for main_sec in root:
            for sub_sec in main_sec:
                print(main_sec.tag + '_' + sub_sec.tag + ': ' + str(sub_sec.text))
        return True
    else:
        return False

if (len(sys.argv) != 2):
    print("Usage: exract.py string:path_of_files_to_check\n")
    raise SystemExit

#mypath = 'P:\\Temp\\Tapio_Develop\\receipt\\'
mypath = sys.argv[1]
mypath = mypath.replace('\\\\', '\\')

onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

for f in onlyfiles:
    extractAttachments(mypath, f)