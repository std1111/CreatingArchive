import os
import zipfile
import datetime
import argparse
import sys
import time
from  time import sleep
import natsort


import pickle
from googleapiclient.discovery import build
from googleapiclient import errors
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient.http import MediaFileUpload



import base64
import smtplib
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def create_message(sender, to, subject, message_text):
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}




def send_message(servicemail,to_email, user_id, success,subject):
    if success:
        subject = 'Успешная загрузка файла на Google Drive %s' %subject
    else:
        subject = 'Внимание! Ошибка загрузки файла на Google Drive %s' %subject
    message_text = subject

    message = create_message('triumph.kiev@gmail.com', to_email, subject, message_text) 
    try:
        message = (servicemail.users().messages().send(userId=user_id, body=message)
               .execute())
        print ('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print ('An error occurred: %s' % error)



def add_info_to_journal(info):
    with open(OUTPUT_FILE, "a+") as file_object:
        file_object.seek(0)
        data = file_object.read(100)
        if len(data) > 0:
            file_object.write("\n")
            text_to_append =time.strftime("%Y-%m-%d-%H.%M.%S", time.localtime()) + ' ' + str(info)
            print(text_to_append)
            file_object.write(text_to_append)
            file_object.close()





# Creating service
def createService(api,version,_path,_scopes):
    creds = None  
    path_token = '%s/token.pickle' %(_path)
    credentials = '%s/credentials.json' %(_path)
    print(path_token)
    _service=''
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(path_token):
        with open(path_token, 'rb') as token:
            creds = pickle.load(token)
       
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials, _scopes)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(path_token, 'wb') as token:
            pickle.dump(creds, token)
    _service = build(api,version, credentials=creds)
    return _service




def createParser ():
    parser = argparse.ArgumentParser()
    parser.add_argument ('-BASE_DIR','--BASE_DIR',default='Z:\\1C_Base\\WEB100') 
    parser.add_argument ('-ARCH_DIR','--ARCH_DIR',default='Z:\\Backup\\Conf_1C7\WEB100')
    parser.add_argument ('-PREFIX_ARCH','--PREFiX_FILE_ARCH',default= 'WEB100')
    parser.add_argument ('-FOLDERID','--DRIVE_FOLDER_ID',default= '1xSWxsCCKhLy-JysoL8uslnXyjIVvFitC')
    parser.add_argument ('-TEST','--TESTMODE',default='True')
    return parser

def getSize(filename):
    #print (filename)
    st = os.stat(filename)
    return st.st_size





BASE_DIR = ''
ARCH_DIR  = ''
PREFiX_FILE_ARCH = ''
LABELTIME =  datetime.datetime.now().strftime("_%Y-%m-%d_%H%M%S")
TESTMODE = True
DRIVE_FOLDER_ID = ''
OUTPUT_FILE = 'journal.log'
EMAIL = 'stdmomo7@gmail.com'
NUMBER_COPIES = 0


MAIN_FOLDER = 'Z:\\Backup\\CreatingArchive\\'
FOLDER_WITH_DRIVE_TOKEN = MAIN_FOLDER + 'Drive_Token'
SCOPES_GOOGLE_DRIVE_GMAIL  = ['https://www.googleapis.com/auth/drive.metadata.readonly','https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/gmail.send']




def main():
    add_info_to_journal('------------------------------ ')
    add_info_to_journal('START')
    name_file = PREFiX_FILE_ARCH+ LABELTIME +'.zip'
    full_path_to_file =  ARCH_DIR +'\\' + name_file
    backup_zip = zipfile.ZipFile(full_path_to_file, 'w')
    add_info_to_journal('Creating %s' % name_file)
    for folder, subfolders, files in os.walk(BASE_DIR):
        for file in files:
            if not file.endswith('.CDX'):
                backup_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder,file), BASE_DIR), compress_type = zipfile.ZIP_DEFLATED) 
    backup_zip.close()
    add_info_to_journal('Done')
    # Compare with size of last copy
    archive_files = [] #list of zip filenames
    dirFiles = os.listdir(ARCH_DIR) #list of directory files
    dirFiles = sorted(dirFiles,reverse=True) #sort numerically in descending order
    for files in dirFiles: #filter out all non zip
        if ('.zip' in files) & (PREFiX_FILE_ARCH in files):
             archive_files.append(files)    
    if (len(archive_files)>1):
        print(archive_files[1])
        size_last_archive =  getSize(ARCH_DIR +'\\' + archive_files[1])
        size_cur_archive  =  getSize(full_path_to_file)
        if (size_last_archive == size_cur_archive):
            if os.path.exists(full_path_to_file):
                os.remove(full_path_to_file)
                add_info_to_journal('Sizes of last copy and current one are equal')
                add_info_to_journal('Deleting %s' %name_file)
                add_info_to_journal('End processing')
                return
        #Check number of archive copies    
        num_exc_copies = len(archive_files) - NUMBER_COPIES
        if  num_exc_copies > 0:
            for i in range(num_exc_copies):
                archive = ARCH_DIR +'\\' + archive_files[len(archive_files)-i-1]
                if os.path.exists(archive):
                    os.remove(archive)
                    add_info_to_journal('Deleting exceeding copy %s' %archive)      
    #Download to drive
    print('Begin upload to Google Drive %s' % full_path_to_file)                 
    add_info_to_journal('Begin upload to Google Drive %s' % full_path_to_file)   
    try:
        #-------- Upload archive to Google Drive ---------------------------------
        serviceDrive    = createService('drive','v3',FOLDER_WITH_DRIVE_TOKEN,SCOPES_GOOGLE_DRIVE_GMAIL)
        file_metadata = {'name': name_file,'parents': [DRIVE_FOLDER_ID]}
        archive = MediaFileUpload(full_path_to_file,resumable=True)
        _file = serviceDrive.files().create(body=file_metadata,
                                        media_body=archive,
                                        fields='id').execute()
        add_info_to_journal('File ID on Drive: %s' % _file.get('id'))
        add_info_to_journal('End upload')
        #serviceGMAIL    = createService('gmail', 'v1',FOLDER_WITH_DRIVE_TOKEN,SCOPES_GOOGLE_DRIVE_GMAIL)
        #send_message(serviceGMAIL,EMAIL, "me", True, full_path_to_file)
        
        # Удаление неактуальных архивных копий
        add_info_to_journal('Removing unnecessary copies from Google Drive') 
        max_limit_backup_files = 5
        filter = "'" + DRIVE_FOLDER_ID + "' in parents "
        filter = filter  + "and name contains '.zip' and name contains '"+ PREFiX_FILE_ARCH + "' "
        results = serviceDrive.files().list(q= filter , spaces='drive',
            pageSize=100, orderBy="createdTime,modifiedTime ",fields="nextPageToken, files(id, name, createdTime)").execute()
        items = results.get('files', [])
        count_files = 0
        add_info_to_journal('Number files in folder : %s' % (len(items)))
        if not items:
            add_info_to_journal('No files found.')
        else:
            add_info_to_journal('List of files in folder:')
            for item in items:
                add_info_to_journal(u'{0} ({1}) created time: {2}'.format(item['name'], item['id'], item['createdTime']))
                count_files +=1
        numberFilesForDelete = count_files - max_limit_backup_files 
        add_info_to_journal("Count files for deleting: %s" %numberFilesForDelete)
        if (numberFilesForDelete > 0):
            i = 0
            for item in items:
                i+=1
                if (i <= numberFilesForDelete):
                    add_info_to_journal('Deleting file: {0} id : {1} created time: {2}'.format(item['name'],item['id'], item['createdTime']))
                    file_id = item['id'] 
                    print(file_id)
                    serviceDrive.files().delete(fileId = file_id).execute()    
                else:
                    break
        
    except errors.HttpError as error:
        serviceGMAIL    = createService('gmail', 'v1',FOLDER_WITH_DRIVE_TOKEN,SCOPES_GOOGLE_DRIVE_GMAIL)
        send_message(serviceGMAIL,EMAIL, "me", False, full_path_to_file)
        add_info_to_journal('An error occurred: %s' % error)
        add_info_to_journal('Error uploading to Google Drive %s' % full_path_to_file)

    



if __name__ == '__main__':
    
    parser = createParser()
    namespace = parser.parse_args(sys.argv[1:])

    print (namespace)
    BASE_DIR =  namespace.BASE_DIR
    ARCH_DIR =  namespace.ARCH_DIR 
    PREFiX_FILE_ARCH = namespace.PREFiX_FILE_ARCH
    DRIVE_FOLDER_ID  = namespace.DRIVE_FOLDER_ID
    OUTPUT_FILE = ARCH_DIR+'\\journal.log'
    NUMBER_COPIES = 5

    TESTMODE = namespace.TESTMODE
    if TESTMODE== 'False':
        TESTMODE = False
    else:
        TESTMODE = True
        
    print('Arguments ' )
    print('BASE_DIR: ' + '%s' % (BASE_DIR))
    print('ARCH_DIR: '+ '%s' % (ARCH_DIR))
    print('PREFiX_FILE_ARCH :'+ '%s' % (PREFiX_FILE_ARCH))
    print('DRIVE_FOLDER_ID :'+ '%s' % (DRIVE_FOLDER_ID))
    sleep(3)


    main()




## or file.endswith('.pdf') or file.endswith('.zip'))
