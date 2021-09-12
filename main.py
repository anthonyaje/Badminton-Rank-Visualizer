from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
from pyasn1.type.univ import Null
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import pdb
print("Google Sheets badminton rank visualizer by anthonyaje")

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
modestGrassRatingPointsSheetId = "1GgW0t9RgVqwZXQ7kZBdYQWuE4ahveN5StkgeGFnI_Cw"

# Using my credentials.json try to authenticate with Google API
def ServiceAuth():
    print("--- Start authenticating... ---")
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("--- Building service... ---")
    service = build('sheets', 'v4', credentials=creds)
    return service

# Every week the organizer will craete a new 'Sheet' (tab) with the updated score and rank
# This function pull from the last `numDataPoints` sheets from the most recent Sheet
# TODO: validation if the newly added Sheet has different schema / layout
def GetSheetsRange(service, numDataPoints = 10):
    sheetMetadata = service.spreadsheets().get(spreadsheetId=modestGrassRatingPointsSheetId).execute()
    sheets = sheetMetadata.get('sheets', '')
    sheetRanges = []
    for i in reversed(range(numDataPoints)):
        title = sheets[i].get("properties", {}).get("title")
        sheetRanges.append(title)
    return sheetRanges

# From the `ranges` of Sheets, pull up the data of the person of interest.
def GetPlayerData(service, ranges, fullname):
    # Call the Sheets API
    sheet = service.spreadsheets()         
    result = sheet.values().batchGet(spreadsheetId=modestGrassRatingPointsSheetId, ranges=ranges).execute()
    valueRanges = result.get('valueRanges', [])
    personData = []
    for values in valueRanges:
        if not values['values']:
            print('Sheet %s not found.' % values['range'])
        else:
            print('Sheet %s found.' % values['range'])
            personData.append([x for x in values['values'] if len(x) > 2  and x[1] == fullname])
    return personData

def PlotResult():
    plt.gca().invert_yaxis()
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
    plt.xlabel('Date')
    plt.ylabel('Rank')
    plt.title('Badminton rank visualization')
    plt.show()

if __name__ == '__main__':
    print("--- Start ---")
    players = ['Anthony Aje', 'Ardiles Setiadi', 'Nico Karim', 'Claire Gunawan']
    service = ServiceAuth()
    ranges = GetSheetsRange(service, 10)

    # TODO
    #   - improve performance for multiplayer by picking everyone data from each open Sheet at once instead of doing player by player.
    for player  in players: 
        playerData = GetPlayerData(service, ranges, player)
        playerRank = [int(x[0][0]) if x else None for x in playerData]
        plt.plot(ranges, playerRank, label=player)
        plt.legend(loc="lower left")

    PlotResult()    

    print("--- Done ---")