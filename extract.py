import PyPDF2
from os import listdir
from os.path import isfile, join
from os import walk
import sys

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
    handler = open(path + filename, 'rb')
    reader = PyPDF2.PdfFileReader(handler)
    dictionary = getAttachments(reader)
    #print(dictionary)
    for fName, fData in dictionary.items():
        with open(path + filename[:-4] + '_' + fName, 'wb') as outfile:
            outfile.write(fData)


if (len(sys.argv) != 2):
    print("Usage: exract.py string:path_of_files_to_check\n")
    raise SystemExit

#mypath = 'P:\\Temp\\Tapio_Develop\\receipt\\'
mypath = sys.argv[1]
mypath = mypath.replace('\\\\', '\\')

onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]



for f in onlyfiles:
    extractAttachments(mypath, f)
    