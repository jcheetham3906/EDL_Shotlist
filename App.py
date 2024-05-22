from flask import Flask, request, render_template, redirect, url_for
import os
import re
import cv2
import time
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']

def authenticate_gdrive_and_sheets(credentials_path):
    creds = None
    token_path = 'token.pickle'
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    return drive_service, sheets_service

def read_edl(edl_content):
    clip_info = []
    current_clip_name = None
    current_in_point = None

    lines = edl_content.splitlines()
    for line in lines:
        if line.startswith("*FROM CLIP NAME:"):
            current_clip_name = line.split(":")[-1].strip()
        else:
            match = re.match(r'\d+\s+\S+\s+V\s+C\s+\d+:\d+:\d+:\d+\s+\d+:\d+:\d+:\d+\s+(\d+:\d+:\d+:\d+)\s+\d+:\d+:\d+:\d+', line)
            if match:
                current_in_point = timecode_to_seconds(match.group(1).strip())

        if current_clip_name and current_in_point is not None:
            clip_info.append((current_clip_name, current_in_point))
            current_clip_name = None
            current_in_point = None

    return clip_info

def timecode_to_seconds(timecode):
    hours, minutes, seconds, frames = map(int, timecode.split(':'))
    return hours * 3600 + minutes * 60 + seconds + frames / 24

def capture_screenshot(video_path, output_path, time_in_seconds):
    vidcap = cv2.VideoCapture(video_path)
    vidcap.set(cv2.CAP_PROP_POS_MSEC, time_in_seconds * 1000)
    success, image = vidcap.read()
    if success:
        cv2.imwrite(output_path, image)
    else:
        print(f"Failed to capture screenshot at {time_in_seconds} seconds")
    vidcap.release()

def upload_image(service, file_path):
    file_metadata = {'name': os.path.basename(file_path)}
    media = MediaFileUpload(file_path, mimetype='image/jpeg')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = file.get('id')
    service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()
    return file_id

def generate_shareable_link(file_id):
    link = f"https://drive.google.com/uc?export=view&id={file_id}"
    return link

def create_new_sheet(service, spreadsheet_id, sheet_title):
    body = {
        'requests': [{
            'addSheet': {
                'properties': {
                    'title': sheet_title
                }
            }
        }]
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
    return sheet_id

def update_google_sheet(service, spreadsheet_id, sheet_title, data):
    sheet_id = create_new_sheet(service, spreadsheet_id, sheet_title)
    sheet = service.spreadsheets()
    body = {
        'values': data
    }
    result = sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=f'{sheet_title}!A1',
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    if len(data) > 1:
        requests = []
        for i in range(1, len(data)):
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": i,
                        "endIndex": i + 1
                    },
                    "properties": {
                        "pixelSize": 100
                    },
                    "fields": "pixelSize"
                }
            })

        body = {
            'requests': requests
        }
        sheet.batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        edl_file = request.files['edl_file']
        video_file = request.files['video_file']
        sheet_title = request.form['sheet_title']
        credentials_file = request.files['credentials_file']

        if edl_file and video_file and sheet_title and credentials_file:
            edl_content = edl_file.read().decode("utf-8")
            video_path = os.path.join('uploads', video_file.filename)
            credentials_path = os.path.join('uploads', credentials_file.filename)
            
            os.makedirs('uploads', exist_ok=True)
            
            video_file.save(video_path)
            credentials_file.save(credentials_path)
            
            clip_info = read_edl(edl_content)
            drive_service, sheets_service = authenticate_gdrive_and_sheets(credentials_path)
            
            spreadsheet_id = '1wOSpDxXF_K7AkksRAl77N_4MZ9FJxm0RsOWI3uxOvk8'
            data = [['Image', 'Filename']]
            
            for idx, (clip_name, in_point) in enumerate(clip_info):
                output_image_path = os.path.join('uploads', f"screengrab_{idx}.jpg")
                capture_screenshot(video_path, output_image_path, in_point)
                
                if os.path.exists(output_image_path):
                    file_id = upload_image(drive_service, output_image_path)
                    time.sleep(1)
                    image_url = generate_shareable_link(file_id)
                    data.append([f'=IMAGE("{image_url}")', clip_name])
            
            update_google_sheet(sheets_service, spreadsheet_id, sheet_title, data)
            return redirect(url_for('success', sheet_id=spreadsheet_id))
    
    return render_template('index.html')

@app.route('/success')
def success():
    sheet_id = request.args.get('sheet_id')
    return f"Google Sheet updated: <a href='https://docs.google.com/spreadsheets/d/{sheet_id}'>here</a>"

if __name__ == "__main__":
    app.run(debug=True)
