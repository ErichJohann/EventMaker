# Google Calendar Google Sheets Event Importer
This script automatically creates Google Calendar events based on a Google Sheets file.

### Features:
- OAuth2 authentication with Google
- Token storage and refresh
- Read multiple events from a Google Sheets document
- Usage via command line with either sheet name or URL
- Supports both sheet name and direct spreadsheet link
- Validates date and time formats before creating events
---

### Requirements
- Google Cloud Console set up
- Install required libraries:
  - gspread
  - google-auth
  - google-auth-oauthlib
  - google-api-python-client
- Google API credentials (`client_secret.json` file)

---

### Google Cloud
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Activate **Google Sheets API** and **Google Calendar API**
4. Necessary scopes:
   - `https://www.googleapis.com/auth/calendar`
   - `https://www.googleapis.com/auth/spreadsheets`
   - `https://www.googleapis.com/auth/drive`
5. Create OAuth 2.0 Client ID credentials
6. Download the `.json` credentials file and place it in the project folder
7. The account used must be added as a beta tester for unpublished projects (if applicable)

---

### Google Sheets File
- Contains all events which will be created
- Must contain these columns:
  - `event` — Event title
  - `description` — Event details
  - `location` — Event location
  - `start` — Event start date/time
  - `end` — Event end date/time
  - `color` — Google Calendar color ID

---

### Running
```bash python eventmaker.py [Sheet_Name or Sheet_URL]```

On the first run, a browser window will open for Google authentication.
Afterward, a token is saved in token.pickle for reuse.