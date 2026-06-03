#!/usr/bin/env python3
"""
Family Driving School — Enrollment System Setup
------------------------------------------------
Run this once. It will:
  1. Authenticate with your Google account (opens browser)
  2. Create the 'Enrollments' Google Sheet with formatted headers
  3. Create & deploy the Apps Script web app endpoint
  4. Automatically write the URL into index.html

BEFORE RUNNING:
  You need a credentials.json file in this folder.
  Follow these steps (takes ~2 minutes):

  1. Go to https://console.cloud.google.com
  2. Create a new project (or select an existing one)
  3. Search "APIs & Services" → Library
  4. Enable: "Google Sheets API"  AND  "Apps Script API"
  5. Go to APIs & Services → Credentials
  6. Click "Create Credentials" → OAuth client ID
  7. Application type: Desktop app  →  Name it anything  →  Create
  8. Click the download icon  →  save as 'credentials.json' in THIS folder

  ALSO: go to https://script.google.com/home/usersettings
  and make sure "Google Apps Script API" is turned ON.
"""

import json
import os
import sys
import re

# ── Auto-install deps if missing ─────────────────────────────
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    print("Installing required packages…")
    os.system(f'{sys.executable} -m pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client -q')
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

# ── Config ───────────────────────────────────────────────────
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.deployments',
    'https://www.googleapis.com/auth/drive',
]

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
CREDS_FILE   = os.path.join(SCRIPT_DIR, 'credentials.json')
TOKEN_FILE   = os.path.join(SCRIPT_DIR, '.token.json')
HTML_FILE    = os.path.join(SCRIPT_DIR, 'index.html')

SHEET_HEADERS = [
    'Timestamp', 'First Name', 'Last Name', 'Date of Birth',
    'Email', 'Phone', 'Address',
    'Course', 'Preferred Start', 'Schedule',
    'Has Permit', 'Permit #', 'Experience',
    'Emergency Name', 'Emergency Relation', 'Emergency Phone',
    'Preferred Language', 'How Heard', 'Comments',
]

# Navy color for header row: #0C1F3F → {r:0.047, g:0.122, b:0.247}
NAVY = {'red': 0.047, 'green': 0.122, 'blue': 0.247}
GOLD = {'red': 0.961, 'green': 0.694, 'blue': 0.125}
WHITE = {'red': 1.0, 'green': 1.0, 'blue': 1.0}

APPS_SCRIPT_CODE = """
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Enrollments');
  var d = e.parameter;

  sheet.appendRow([
    d.timestamp || new Date().toLocaleString('en-US', {timeZone: 'America/New_York'}),
    d.firstName,  d.lastName,   d.dob,
    d.email,      d.phone,      d.address,
    d.course,     d.startDate,  d.schedule,
    d.hasPermit,  d.permitNumber, d.experience,
    d.emergencyName, d.emergencyRelation, d.emergencyPhone,
    d.language,   d.howHeard,   d.comments
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({success: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('Family Driving School Enrollment API — OK');
}
"""

APPSSCRIPT_MANIFEST = json.dumps({
    "timeZone": "America/New_York",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
        "executeAs": "USER_DEPLOYING",
        "access": "ANYONE_ANONYMOUS"
    }
}, indent=2)


# ── Authentication ───────────────────────────────────────────
def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                print('\n❌  credentials.json not found in this folder.')
                print('    Follow the instructions at the top of this script.\n')
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds


# ── Google Sheet ─────────────────────────────────────────────
def create_spreadsheet(svc):
    print('  Creating Google Sheet…')
    ss = svc.spreadsheets().create(body={
        'properties': {'title': 'Family Driving School — Enrollments'},
        'sheets': [{'properties': {
            'title': 'Enrollments',
            'gridProperties': {'frozenRowCount': 1}
        }}]
    }).execute()
    sid = ss['spreadsheetId']
    sheet_id = ss['sheets'][0]['properties']['sheetId']

    # Write headers
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range='Enrollments!A1',
        valueInputOption='RAW',
        body={'values': [SHEET_HEADERS]}
    ).execute()

    # Format: navy header row, auto-resize columns, alternating row colors
    n = len(SHEET_HEADERS)
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': [
        # Header background + bold white text
        {'repeatCell': {
            'range': {'sheetId': sheet_id, 'startRowIndex': 0, 'endRowIndex': 1,
                      'startColumnIndex': 0, 'endColumnIndex': n},
            'cell': {'userEnteredFormat': {
                'backgroundColor': NAVY,
                'textFormat': {'bold': True, 'foregroundColor': WHITE, 'fontSize': 10},
                'horizontalAlignment': 'CENTER',
                'verticalAlignment': 'MIDDLE',
            }},
            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
        }},
        # Gold bottom border on header
        {'updateBorders': {
            'range': {'sheetId': sheet_id, 'startRowIndex': 0, 'endRowIndex': 1,
                      'startColumnIndex': 0, 'endColumnIndex': n},
            'bottom': {'style': 'SOLID_MEDIUM', 'color': GOLD}
        }},
        # Auto-resize all columns
        {'autoResizeDimensions': {
            'dimensions': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                           'startIndex': 0, 'endIndex': n}
        }},
        # Set row height for header
        {'updateDimensionProperties': {
            'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                      'startIndex': 0, 'endIndex': 1},
            'properties': {'pixelSize': 36},
            'fields': 'pixelSize'
        }},
    ]}).execute()

    url = f'https://docs.google.com/spreadsheets/d/{sid}'
    print(f'  ✓ Sheet ready: {url}')
    return sid


# ── Apps Script ──────────────────────────────────────────────
def deploy_script(script_svc, spreadsheet_id):
    print('  Creating Apps Script project…')
    project = script_svc.projects().create(body={
        'title': 'FDS Enrollment API',
        'parentId': spreadsheet_id,
    }).execute()
    script_id = project['scriptId']

    # Upload code + manifest
    script_svc.projects().updateContent(
        scriptId=script_id,
        body={'files': [
            {'name': 'Code',        'type': 'SERVER_JS', 'source': APPS_SCRIPT_CODE},
            {'name': 'appsscript',  'type': 'JSON',      'source': APPSSCRIPT_MANIFEST},
        ]}
    ).execute()
    print('  ✓ Script uploaded')

    # Create a versioned snapshot
    version = script_svc.projects().versions().create(
        scriptId=script_id,
        body={'description': 'v1 – initial enrollment endpoint'}
    ).execute()

    # Deploy as web app
    print('  Deploying web app…')
    deployment = script_svc.projects().deployments().create(
        scriptId=script_id,
        body={
            'versionNumber': version['versionNumber'],
            'manifestFileName': 'appsscript',
            'description': 'Enrollment form endpoint',
        }
    ).execute()

    deployment_id = deployment['deploymentId']
    url = f'https://script.google.com/macros/s/{deployment_id}/exec'
    print(f'  ✓ Deployed: {url}')
    return url


# ── Patch index.html ─────────────────────────────────────────
def patch_html(web_app_url):
    if not os.path.exists(HTML_FILE):
        return False
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        content = f.read()
    updated = content.replace(
        "'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE'",
        f"'{web_app_url}'"
    )
    if updated == content:
        return False
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(updated)
    return True


# ── Main ─────────────────────────────────────────────────────
def main():
    print()
    print('=' * 58)
    print('  Family Driving School — Enrollment Setup')
    print('=' * 58)
    print()

    print('Step 1/3  Authenticating with Google…')
    creds = authenticate()
    print('  ✓ Authenticated\n')

    print('Step 2/3  Setting up Google Sheet…')
    sheets_svc = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = create_spreadsheet(sheets_svc)
    print()

    print('Step 3/3  Deploying Apps Script endpoint…')
    script_svc = build('script', 'v1', credentials=creds)
    try:
        web_app_url = deploy_script(script_svc, spreadsheet_id)
    except HttpError as e:
        body = json.loads(e.content)
        msg = body.get('error', {}).get('message', str(e))
        if 'Apps Script API' in msg or 'script' in msg.lower():
            print(f'\n  ⚠️  Apps Script API not enabled for your account.')
            print('     Go to: https://script.google.com/home/usersettings')
            print('     Turn ON "Google Apps Script API", then re-run this script.\n')
        else:
            print(f'\n  ❌  Deployment error: {msg}\n')
        sys.exit(1)

    print()
    patched = patch_html(web_app_url)

    print('=' * 58)
    print('  ✅  All done!')
    print('=' * 58)
    print()
    print(f'  Sheet   → https://docs.google.com/spreadsheets/d/{spreadsheet_id}')
    print(f'  API URL → {web_app_url}')
    if patched:
        print()
        print('  ✓ URL written into index.html automatically.')
        print('    Your enrollment form is fully connected.')
    else:
        print()
        print('  ⚠️  Could not auto-patch index.html.')
        print(f"     Open index.html and replace YOUR_GOOGLE_APPS_SCRIPT_URL_HERE")
        print(f"     with: {web_app_url}")
    print()


if __name__ == '__main__':
    main()
