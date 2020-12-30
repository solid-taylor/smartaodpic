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
from xlwt import Workbook, Formula
import sqlite3 

""" TODO: import this module, otherwise it's not safe
import defusedxml
import defusedxml.ElementTree as ET """

#default fields specificated by Hungarian Post Office & source-user fields
colnames={
'efj_adatok_efj_zaras' :  0,
'efj_adatok_efj_szoftver' :  0,
'efj_adatok_xsd_verzio' :  0,
'felado_felado_megallapodas' :  0,
'felado_felado_nev' :  0,
'felado_felado_irsz' :  0,
'felado_felado_hely' :  0,
'felado_felado_kozterulet_nev' :  0,
'felado_felado_kozterulet_jelleg' :  0,
'felado_felado_hazszam' :  0,
'felado_felado_kozelebbi_cim' :  0,
'felado_felado_epulet' :  0,
'felado_felado_lepcsohaz' :  0,
'felado_felado_emelet' :  0,
'felado_felado_ajto' :  0,
'felado_felado_postafiok' :  0,
'kuldemeny_azonosito' :  0,
'atvetel_idopont' :  0,
'atvetel_atvevo_nev' :  0,
'atvetel_atvetel_jogcim' :  0,
'atvetel_visszakuldes_oka' :  0,
'kuldemeny_tv_sajat_jelzes' :  0,
'kuldemeny_felvetel_datum' :  0,
'kuldemeny_cimzett_nev' :  0,
'kuldemeny_cimzett_irsz' :  0,
'kuldemeny_cimzett_hely' :  0,
'kuldemeny_cimzett_kozterulet_nev' :  0,
'kuldemeny_cimzett_kozterulet_jelleg' :  0,
'kuldemeny_cimzett_hazszam' :  0,
'kuldemeny_cimzett_kozelebbi_cim' :  0,
'kuldemeny_cimzett_epulet' :  0,
'kuldemeny_cimzett_lepcsohaz' :  0,
'kuldemeny_cimzett_emelet' :  0,
'kuldemeny_cimzett_ajto' :  0,
'kuldemeny_cimzett_postafiok' :  0,
'kuldemeny_sajat_azonosito' :  0,
'kuldemeny_tv_vonalkod' :  0,
'kuldemeny_tv_vonalkod_tipus' :  0,
'kuldemeny_hiv_iratszam' :  0,
'kuldemeny_hiv_irat_fajta' :  0,
'kuldemeny_hiv_ertesito' :  0,
'file_id': 0,
'file_source' : 0,
'userid' : 0,
'sessionid' : 0
}

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

def extractAttachment(path, filename, oTbl, cleanup_after, userid = '0', sessionid = '0'):
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
                    aRecord['userid'] = userid
                    aRecord['sessionid'] = sessionid
                    oTbl.append(aRecord)
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
    sheet1 = wb.add_sheet("kezb_"+ datetime.now().strftime("%m%d%y%H%M%S"), cell_overwrite_ok=True) 
    sheet1.panes_frozen = True
    sheet1.remove_splits = True
    
    sheet1.horz_split_pos = 5
    sheet1.horz_split_first_visible = 5

    row = 0 
    col = 0
    bold = xlwt.easyxf('font: bold 1') 
    sheet1.write(row, col, 'Data extraction from Advice of Delivery Slips of the Hungarian Post. Free to use on "as is" basis!', bold)
    row += 2
    boldred = xlwt.easyxf('font: bold 1, color red;') 
    #         row_start,row_end,col_start,col_end
    sheet1.write_merge(2, 2, 0, 8, Formula('"You need more complex solution? Send a mail to "& HYPERLINK("mailto:develop@vipexkft.hu";"develop@vipexkft.hu") & "!"'), boldred)

    row += 2
    header = xlwt.easyxf('pattern: pattern solid, fore_colour gray40; font: bold 1, color white; borders: left thin, right thin, top thin, bottom thin')
    for aCol in iColnames:
        sheet1.write(row, col, aCol, header)
        col += 1
    with open(oFilename, 'w') as outfile:
        for aRecord in iTbl:
            col = 0
            row += 1
            for aCol in iColnames.keys():
                if (aCol in aRecord.keys()):
                    sheet1.write(row, col, aRecord[aCol].replace('None', ''))
                    col += 1
                else:
                    col += 1
    wb.save(oFilename) 
    return True


#------------------------------------------------    main code -----------------------------------------
if (len(sys.argv) != 2):
    mypath, myfilename = os.path.split(os.path.abspath(__file__))
    if sys.platform.startswith("win32") or sys.platform.startswith("cygwin"):
        mypath = mypath + '\\'
    else:
        mypath = mypath + '/'
else:
    mypath = sys.argv[1]
    mypath = mypath.replace('\\\\', '\\')

c=0
tbl = []

onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
for f in onlyfiles:
    c += extractAttachment(mypath, f, tbl, True)  # set the last param to False if you want the see the temporary xml files
tbl  = normalize_table(tbl,colnames)
print('Total of ' + str(c) +' .pdf files handled in the directory: ' )

#success = table2csv(tbl,colnames, mypath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".csv")
success = table2xls(tbl,colnames, mypath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".xls")

##TODO: get the cleanupafter prop as (command-line) arg
##TODO: if same-name column exists avoid collision (make a new colname "old2")
##TODO: check if there is more than one notification info --> how it appears in the xml???

