from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v3', http=http, discoveryServiceUrl=discoveryUrl)
    print('-'*20)
    print('Enter the ID of the spreadsheet you would like to edit.')
    spreadsheetId = input('Enter "new" to create one or "trial" to use a default spreadsheet. ID: ')

    if spreadsheetId.lower() == "new":
        create(service)

    else:
        view(service, spreadsheetId)
        edit(service, spreadsheetId)


def create(service):
    '''Helps in creating a new spreadsheet for a user.
    Takes title of spreadsheet from user to create a spreadsheet.
    Prints spreadsheetID of the spreadsheet created which is necessary for further use of this spreadsheet.
    '''

    data = {
      "properties": {
        'title': input('\nEnter Title of Spreadsheet: ')
      },
    }
    try:
        newsheet = service.spreadsheets().create(body=data).execute()
        print('Spreadsheet Created! Kindly note the spreadsheet ID for further use.')
        print('Spreadsheet ID:', newsheet['spreadsheetId'])
    except:
        print('Spreadsheet Not Created!')

def view(service, spreadsheetId):
    '''Helps in viewing data for a spreadsheet using spreadsheet ID.

    Takes input for range of data to be viewed using A1 notation- which is described in output.
    Request sent returns a json object consisting of data properties.
    The 'values' key corresponds to a list of actual spreadsheet values which is parsed and printed in terminal.
    
    For demonstration purposes, a "trial" spreadsheet is initiated and used for convenience.
    The spreadsheet ID for trial spreadsheet is hardcoded. It can be accessed at:
    https://docs.google.com/spreadsheets/d/1Uk7T78T5n-HT0p9Y6ggxVF7RR06Lrn8nnZ32bsMdwhw/edit?usp=sharing
    '''
    
    if spreadsheetId == "trial":
        spreadsheetId = '1Uk7T78T5n-HT0p9Y6ggxVF7RR06Lrn8nnZ32bsMdwhw'
        print('Note: Use sheet name \"Sheet1\" for this spreadsheet.')
    
    print('-'*20)                                #Explaining A1 notation to user
    print('Please note the A1 Notation: This is a string like \"Sheet1!A1:B2\", that refers to a group of cells in the spreadsheet.')
    print('For example, valid ranges are:')
    print('\tSheet1!A1:B2 refers to the \"top-left-cell:bottom-right-cell in Sheet1\". It displays first two cells in the top two rows of Sheet1.')
    print('\tSheet1!A:A refers to all the cells in the first column of Sheet1.')
    print('\tSheet1!1:2 refers to the all the cells in the first two rows of Sheet1.')
    print('\tSheet1!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.')
    print('\tA1:B2 refers to the first two cells in the top two rows of the first visible sheet.')
    print('\tSheet1 refers to all the cells in Sheet1.')
    print('Spreadsheet API uses API notation for accessing spreadsheets. Fill the data keeping that in mind.')

    sheetName = input('\nEnter the name of sheet you would like to see (Enter "$1" to see data of first sheet): ')
    cellRange = input('\nCell Range as per A1 notation (eg: A1:B2) or enter \"all\" to see the entire sheet (sheet name must be explicitly mentioned for this): ')
    
    if ':' not in cellRange and cellRange.lower() != 'all':
        print('Illegal input. Please see A1 Notation!')
        return
    
    if sheetName == "$1" and cellRange.lower() == "all":
        print('At least one of sheet-name or cell-range must be explicitly mentioned. Please try again.\nHint: By default, sheets are named \"Sheet1\", \"Sheet2\", \"Sheet3\" etc in order. If sheet name was not explicitly changed, this convention may be tried.')
        return

    elif sheetName != "$1" and cellRange.lower() != "all":
        rangeName = sheetName+'!'+cellRange

    elif sheetName != "$1" and cellRange.lower() == "all":
        rangeName = sheetName

    elif sheetName == "$1" and cellRange.lower() != "all":
        rangeName = cellRange
    
    print('Fetching data for', rangeName)
    try:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
        print('')
        values = result['values']
        if not values:
            print('No data found.')
        else:
            for row in values:
                curr_row = ''
                for col in row:
                    curr_row = curr_row + str(col) + ' , '
                print(curr_row)
    except:
        print('\nUnable to find requested block of data. Please check the input format for sheet name and cell range.')

def edit(service, spreadsheetId):
    pass

if __name__ == '__main__':
    main()