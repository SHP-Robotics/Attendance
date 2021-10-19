import datetime
import time
from googleapiclient.discovery import build

# From OAuth 2.0 for Service Accounts
from google.oauth2 import service_account
SERVICE_ACCOUNT_FILE = 'credentials.json'                              # Client secrets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']       # Spreasheet: Editor - Read & Write
creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of a sample spreadsheet.
spreadsheet_id = '1Lij9_onCKsr7CWVI4-GKEz73XMwEXljDsUP6aqcJEV0'
att_sheet = 'All Team 9/17'
input_range = att_sheet + '!A1:D100'
output_range = 'All Team!'                    # Open ended
roster_range = 'All Team!A1:X100'              # Open ended
value_input_option = "USER_ENTERED"
service = build('sheets', 'v4', credentials=creds)     # "sheets" api version 4
# Call the Sheets API
sheet = service.spreadsheets()
rawattlists    = sheet.values().get(spreadsheetId=spreadsheet_id, range=input_range).execute()
rawroslists = sheet.values().get(spreadsheetId=spreadsheet_id, range=roster_range).execute()
attlists  = rawattlists.get('values', [])
roslists = rawroslists.get('values', [])

num_att_names = len(attlists)
num_ros_names = len(roslists)
num_cols = len(roslists[0])
new_cols = chr(ord('A') + num_cols)             # New column in alphabet          
todaysDate = datetime.date.today()
request = sheet.values().update(spreadsheetId=spreadsheet_id, range=output_range+new_cols+'1', 
                                        valueInputOption="RAW", body={'values':[[str(todaysDate)]]}).execute()
print ("The new column is: " + new_cols)
request = sheet.values().update(spreadsheetId=spreadsheet_id, range=output_range+new_cols+'1', 
                                    valueInputOption="RAW", body={'values':[[att_sheet]]}).execute()
num_reg = 0
num_came = num_att_names - 1
for i in range (1,num_att_names):               # Checking all the names in the attendance list
    last  = attlists[i][1].replace(" ","")                      # Column 0 is timestamp
    last  = last.lower()
    first = attlists[i][2].replace(" ","")    
    first = first.lower()
    j = 1
    found = False
    # Find the person in the roster
    while (not(found) and (j in range (1, num_ros_names))):
        range_cell = output_range+new_cols+str(j+1)     # Cell Rows starts with 1 not zero
        rlast = roslists[j][0].replace(" ","")
        rlast = rlast.lower()
        if ((last in rlast) or (rlast in last)):             # matched last name
            rfirst = roslists[j][1].replace(" ","")
            rfirst = rfirst.lower()
            rnick  = roslists[j][2].replace(" ","")
            rnick = rnick.lower()
            if ((rfirst in first) or (rnick in first)):       # Found: matching first and last names
                found = True
                num_reg += 1
                request = sheet.values().update(spreadsheetId=spreadsheet_id, range=range_cell, 
                                    valueInputOption="RAW", body={'values':[["x"]]}).execute()
        j += 1
    if (not(found)):
        print ("Intruder: " + attlists[i][2] + ' ' + attlists[i][1])
# Fill the empty cell with A
rawroslists = sheet.values().get(spreadsheetId=spreadsheet_id, range=roster_range).execute()
roslists = rawroslists.get('values', [])
num_cols = len(roslists[0])                         # Updated range
print (num_came, "people came. ", num_reg, "people registered.", (num_came - num_reg), "intruders.")