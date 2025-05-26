import os
import sys
import pickle
from datetime import datetime
import gspread
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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

def main(sheet):
    cred = getCredentials()

    try:
        #Gsheet = build("sheets", "v4", credentials = cred)
        gsheet = gspread.authorize(cred)
        gcalendar = build("calendar", "v3", credentials = cred)

        sh = gsheet.open(sheet)
        print(sh.sheet1.get('A1'))

    except HttpError as e:
        print(e)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python eventmaker.py [Gspread_Sheet_Name]")
    else:
        main(sys.argv[1])