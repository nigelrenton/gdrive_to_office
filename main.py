from __future__ import print_function
import io
import os
import shutil
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.http import MediaFileUpload

# uses a service account with keys in service_key.json

scope = 'https://www.googleapis.com/auth/drive'
key_file_location = 'service_key.json'
file_count = 0

# creates the service function to connect to g drive
def get_service(api_name, api_version, scopes, key_file_location):

    credentials = service_account.Credentials.from_service_account_file(key_file_location)
    scoped_credentials = credentials.with_scopes(scopes)
    service = build(api_name, api_version, credentials=scoped_credentials)
    return service

# checks for duplicates by filename with extension e.g., document.docx
def chkdup(name,parent):

    print('checking for duplicates of ' + name + ' in ' + parent)
    dup_count = 0
    dup_found = []

    try:
        service = get_service(
            api_name='drive',
            api_version='v3',
            scopes=[scope],
            key_file_location=key_file_location)
        chkdup = service.files().list(q="name='" + name + "' and '" + parent + "' in parents", spaces='drive', fields='files(id,parents)').execute()
        for f in chkdup.get('files', []):
            dup_count = dup_count +1
        if dup_count > 0: 
            return(True)
        return(False)
    except:
        return(False)

# deletes the locally cached file from ./convertio
def delete_local_file(name):

    print('deleting local copy of ' + name)
    os.remove('convertio/' + name)

# export (download) as ms office xml, check for duplicates, upload if no duplicate is found
def convert(gfile_id, nfile_mime, gfile_name, gfile_parent, nfile_name):

    global file_count

    try:
        service = get_service(
            api_name='drive',
            api_version='v3',
            scopes=[scope],
            key_file_location=key_file_location)
        
        # download the file in office mimetype
        print('downloading item: ' + str(file_count) + ' id: ' + gfile_id + " name: " + gfile_name)
        dl_req = service.files().export_media(fileId = gfile_id, mimeType=nfile_mime)
        dl_file = io.BytesIO()
        downloader = MediaIoBaseDownload(dl_file, dl_req)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(F'downloaded: {int(status.progress() * 100)}%')
        
        # save to disk
        print('creating: ' + nfile_name)
        dl_file.seek(0)
        with open('convertio/' + nfile_name, 'wb') as f:
                shutil.copyfileobj(dl_file, f, length=524288)
        
        # check for duplicates
        dup_file = chkdup(nfile_name, str(gfile_parent)[2:-2])
        if dup_file == True:
            print('duplicate found, skipping')
            delete_local_file(nfile_name)
            return
        
        # upload ms xml file to original parent folder
        print('uploading: ' + nfile_name + ' to google drive')
        nfile_metadata = {'name': nfile_name, 'parents': gfile_parent}
        ufile = MediaFileUpload('convertio/' + nfile_name, mimetype=nfile_mime)
        gfile_new = service.files().create(body=nfile_metadata, media_body=ufile, fields='id').execute()
        print ('id: %s' % gfile_new.get('id'), 'created')

        # delete local file
        delete_local_file(nfile_name)

    except HttpError as error:

        print(F'An error occurred: {error}')
        dl_file = None
        return

def main():

    global file_count

    try:

        service = get_service(
            api_name='drive',
            api_version='v3',
            scopes=[scope],
            key_file_location=key_file_location)

        files = []
        page_token = None
        gmime = "mimeType contains 'vnd.google-apps."
        n_mime = ""
        gfile_id = ""

        while True:

            # search for files
            response = service.files().list(q=gmime + "document' or " + gmime + "presentation' or " + gmime + "spreadsheet'",
                                            spaces='drive',
                                            fields='nextPageToken, '
                                                   'files(id, name, mimeType, parents)',
                                            pageToken=page_token).execute()            

            for file in response.get('files', []):
                print(F'{file.get("name")}, {file.get("id")}, {file.get("mimeType")}, {file.get("parents")}')
                file_count = file_count + 1
            
            print('do you want to convert ', file_count, 'files?')

            confirm = ''
            while not(confirm == 'y'):
                confirm = input("(y) yes or (n) no: ")
                if confirm == 'n':
                    print('quitting...')
                    exit()
                    
            for file in response.get('files', []):

                gfile_id = file.get("id")
                gfile_name = file.get("name")
                gfile_parent = file.get("parents")

                # set office mimetypes and extensions
                if file.get("mimeType") == 'application/vnd.google-apps.spreadsheet':
                    n_mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    n_ext = '.xlsx'
                elif file.get("mimeType") == 'application/vnd.google-apps.presentation':
                    n_mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    n_ext = '.pptx'
                elif file.get("mimeType") == 'application/vnd.google-apps.document':
                    n_mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    n_ext = '.docx'
                
                nfile_name = gfile_name.replace("/","-") + n_ext

                # do the conversion
                convert(gfile_id, n_mime, gfile_name, gfile_parent, nfile_name)
              
            files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    except HttpError as error:
        print(f'An error occurred: {error}')
        files = None
    
    return files
    print(files[0])
    

if __name__ == '__main__':

    main()
    
