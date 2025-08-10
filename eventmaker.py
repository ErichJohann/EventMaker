import os
import sys
import pickle
from datetime import datetime
import gspread
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import re

SCOPES = [
    "https://www.googleapis.com/auth/calendar", 
    "https://www.googleapis.com/auth/spreadsheets", 
    "https://www.googleapis.com/auth/drive"
    ]
TOKEN_FILE = "token.pickle"
CLIENT_SECRET_FILE = "client_secret.json"

DATE_FORMAT = "%Y-%m-%dT%H:%M:%S%z"

def getCredentials():
    credentials = None
    
    #Token for easier authentication
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            credentials = pickle.load(token)

    #Refresh token
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        #First execution
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            credentials = flow.run_local_server(port=0)

        #saves newest token
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(credentials, token)

    return credentials

def getSheetId(url):
    #looks for id in url
    pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    else:
        return None

def setEvent(ev,cal):
    print(f"Creating event: {ev['summary']}...")
    cal.events().insert(calendarId="primary", body=ev).execute()

def main(sheet):
    cred = getCredentials()

    try:
        #Gsheet = build("sheets", "v4", credentials = cred)
        gsheet = gspread.authorize(cred)
        gcalendar = build("calendar", "v3", credentials = cred)

        #checks whether acess is by name or link
        if sheet.startswith('http'):
            sheetId = getSheetId(sheet)
            if not sheetId:
                print('URL inválida')
                return
            sh = gsheet.open_by_key(sheetId)
        else:
            sh = gsheet.open(sheet)

        sh = sh.sheet1
        records = sh.get_all_records()

        for row in records:
        #sets event to correct format
            try:
                #validates time format and parces to datetime
                startTime = datetime.strptime(row["start"], DATE_FORMAT)
                endTime = datetime.strptime(row["end"], DATE_FORMAT)

                event = { 
                    "summary":row["event"],
                    "description":row["description"],
                    "location":row["location"],
                    "start": {
                        "dateTime":startTime.isoformat()
                    },
                    "end": {
                        "dateTime":endTime.isoformat()
                    },
                    "colorId": row["color"]
                }
                #create event
                setEvent(event,gcalendar)

            except ValueError as e:
                print(f"invalid date time format - {e}")

            except Exception as e:
                print(f"Error creating event - {e}")

    except HttpError as e:
        print(e)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python eventmaker.py [Sheet_Name or URL]")
    else:
        main(sys.argv[1])