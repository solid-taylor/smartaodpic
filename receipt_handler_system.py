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
import werkzeug
from werkzeug.security import check_password_hash, generate_password_hash
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from itsdangerous import URLSafeSerializer, URLSafeTimedSerializer
from globalcons import *


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

#return-receipt handling functions
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
    sheet1.write(row, col, 'Postai kézbesítési igazolások feldolgozása. A program szabadon felhasználható, kizárólag saját felelősségére!', bold)
    row += 2
    boldred = xlwt.easyxf('font: bold 1, color red;') 
    #         row_start,row_end,col_start,col_end
    sheet1.write_merge(2, 2, 0, 8, Formula('"Teljes körű megoldást szeretne? Kérjük, írjon e-mailt a "& HYPERLINK("mailto:develop@vipexkft.hu";"develop@vipexkft.hu") & " címre!"'), boldred)

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

def table2sql(iTbl, iColnames, oFilename):
    #connect to sql conn
    with sqlite3.connect(oFilename) as db:
            try:
                paramstr = ', '.join(['?'] * len(iColnames))
                colnamestr = ', '.join(iColnames)
                qry = 'INSERT INTO receipt ' + f"({colnamestr}) VALUES ({paramstr})"
                aRecVal = []
                cursor = db.cursor()
                for aRec in iTbl:
                    aRecVal.clear()
                    for col in iColnames:
                        aRecVal.append(aRec[col])
                    cursor.execute(qry, aRecVal)
                db.commit()
                return True
            except Exception as e: 
                return False
    #generate qry
    #execute qry
    #cleanup - close conn
    pass

def exctract_receipts(iPath, iTarget, iDb = DEFDB):
    c=0
    tbl = []

    onlyfiles = [f for f in listdir(iPath) if isfile(join(iPath, f))]
    for f in onlyfiles:
        c += extractAttachment(iPath, f, tbl, True)  # set the last param to False if you want the see the temporary xml files
    tbl  = normalize_table(tbl,colnames)
    targets = {
        'csv': table2csv(tbl,colnames, iPath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".csv"),
        'xls': table2xls(tbl,colnames, iPath + "result_"+ datetime.now().strftime("%m%d%y%H%M%S") +".xls"),
        'sql': table2sql(tbl,colnames, iDb)
    }

    success = targets.get(iTarget, False)

    if (success == True):
        return "All went FINE"
    else:
        return "There was an ERROR"

## User related functions
def user_login(iUserMail, iPass, iIPAddress = '127.0.0.0', iDb = DEFDB):
    with sqlite3.connect(iDb) as db:
            for row in db.execute("SELECT id, username, email, hash FROM users WHERE email = ? LIMIT 1", [iUserMail]):
                if check_password_hash(row[3],iPass):
                    qry = '''INSERT INTO logins (
                        userid,
                        ipaddress
                    )
                    VALUES (
                        ?, ?
                    );
                    '''
                    cursor = db.cursor()
                    cursor.execute(qry, [row[0], iIPAddress])
                    db.commit()
                    return row[0]
                else:
                    return False

def user_register(iUserName, iUserMail, iPass, iMailFrom = MAILFROM, iMailFromPass = MAILFROMPASS, iDb = DEFDB):
    hpwd = generate_password_hash(iPass)
    with sqlite3.connect(iDb) as db:
        try:
            qry = '''INSERT INTO users (
                        username,
                        email,
                        hash
                    )
                    VALUES (?, ?, ?)'''
            cursor = db.cursor()
            cursor.execute(qry, [iUserName, iUserMail, hpwd])
            db.commit()
            if user_send_activation_mail(iUserName, iUserMail, iMailFrom, iMailFromPass):
                if user_activate(iUserName, iDb):
                    return True
                else:   
                    return False
            else:
                return False
        except:
            return False

def user_send_activation_mail(iUserName, iUserMail, iMailFrom=MAILFROM, iPwd = MAILFROMPASS, iMailServerSSL = MAILFROM_SERVER, iMailServerPort = MAILFROM_SERVER_PORT):
    try:
        auth_s = URLSafeSerializer("secret key", "auth")
        token = auth_s.dumps({"activate": 1, "uname": iUserName}, salt='activation-mail')
        msg = MIMEMultipart()
        to = [iUserMail]
        msg['From'] = iMailFrom
        msg['To'] = ', '.join(to)
        msg['Subject'] = '[RRS] Account Activation'
        message = f'Dear {iUserName},\n\nIn order to use our services, you need to activate your account. \n\nTo activate your account, go to the following link:\n\nhttp://linkaddress.com/activate/{token}\n\nIf you are NOT intended to activate, or this message seems unfamiliar to you, then you have NOTHING to do with this message.\n\n\nBest Regards!'
        msg.attach(MIMEText(message))
        mailserver = smtplib.SMTP_SSL(iMailServerSSL, iMailServerPort)
        mailserver.ehlo()
        mailserver.login(iMailFrom, iPwd)
        mailserver.sendmail(iMailFrom, to, msg.as_string())
        mailserver.close()
        return True    
    except Exception as e:
        return False  

def user_activate(iToken, iDb = DEFDB):
    auth_s = URLSafeSerializer(SECRET_KEY)
    try:
        data = auth_s.loads(iToken, , salt='activation-mail')
    except:
        return False
    else:
        with sqlite3.connect(iDb) as db:
            try:
                qry = '''UPDATE users
                    SET active = 1
                    WHERE username = ?;
                    '''
                cursor = db.cursor()
                cursor.execute(qry, [data["uname"]])
                db.commit()
                return True
            except:
                return False

def user_remind(iUserName, iUserMail, iMailFrom=MAILFROM, iPwd = MAILFROMPASS, iMailServerSSL = MAILFROM_SERVER, iMailServerPort = MAILFROM_SERVER_PORT):
    try:
        s = URLSafeTimedSerializer(SECRET_KEY)
        token = s.dumps({"reset": 1, "uname": iUserName}, salt='password-change-mail')
        msg = MIMEMultipart()
        to = [iUserMail]
        msg['From'] = iMailFrom
        msg['To'] = ', '.join(to)
        msg['Subject'] = '[RRS] Password Reminder'
        message = f'Dear {iUserName},\n\nThis is a password reminder. \n\nTo change your password go to the following link:\n\nhttp://linkaddress.com/{token}\n\nIf you are NOT intended to change your password, then you have NOTHING to do with this e-mail.\n\n\nBest Regards!'
        msg.attach(MIMEText(message))
        mailserver = smtplib.SMTP_SSL(iMailServerSSL, iMailServerPort)
        mailserver.ehlo()
        mailserver.login(iMailFrom, iPwd)
        mailserver.sendmail(iMailFrom, to, msg.as_string())
        mailserver.close()    
        return True
    except Exception as e:
        return False

def user_change_password(iToken, iNewPass, iDb = DEFDB):
    s = URLSafeTimedSerializer(SECRET_KEY)
    try:
        data = s.loads(iToken, max_age=3600, salt='password-change-mail')
    except:
        return 'Link invalid or expired. Please try password reset again!'
    else:
        if (data["reset"]==1):
            hpwd = generate_password_hash(iNewPass)
            with sqlite3.connect(iDb) as db:
                try:
                    qry = '''UPDATE users
                        SET hash = ?
                        WHERE username = ?;
                        '''
                    cursor = db.cursor()
                    cursor.execute(qry, [hpwd, data["uname"]])
                    db.commit()
                    return True
                except:
                    return False

##TEST

if (len(sys.argv) != 2):
    mypath, myfilename = os.path.split(os.path.abspath(__file__))
    if sys.platform.startswith("win32") or sys.platform.startswith("cygwin"):
        mypath = mypath + '\\'
    else:
        mypath = mypath + '/'
else:
    mypath = sys.argv[1]
    mypath = mypath.replace('\\\\', '\\')

# TEST SCRIPTS
##print(exctract_receipts(mypath,'xls'))
#success = user_register('testnyuszi5',TESTMAILTO,'pass', MAILFROM, MAILFROMPASS)
#login = user_login(TESTMAILTO,'pass')
#remind = user_remind('testnyuszi5',TESTMAILTO, MAILFROM, MAILFROMPASS)
#send_activ = user_send_activation_mail('testnyuszi5',TESTMAILTO, MAILFROM, MAILFROMPASS)
#activate= user_activate('eyJhY3RpdmF0ZSI6MSwidW5hbWUiOiJ0ZXN0bnl1c3ppNSJ9.b4ul1Jv7MWcefyxMEBmkNlX6dbg')
#user_change_password('eyJyZXNldCI6MSwidW5hbWUiOiJ0ZXN0bnl1c3ppNSJ9.X9C3UA.PldIDZdlv-9H1SShpdHmyLfl7zY','newpass')




##TODO-CHK IF RDY Handle exeptions on password change ie:link expired, bad hash ect.
##TODO globalcons.py-ba visszaírni az értékeket, amiket a .git kitörölt 

##TODO: validate the files BEFORE try to store --> if allready stored - inform the user that it has allready been stored before
##TODO: get the cleanupafter prop as (command-line) arg
##TODO: if same-name column exists avoid collision (make a new colname "old2")
##TODO: check if there is more than one notification info --> how it appears in the xml???
##TODO: reminder link generation with hash

