from __future__ import print_function
import httplib2
import os
import sys

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
    service = discovery.build('sheets', 'v3', http=http, discoveryServiceUrl=discoveryUrl)          #REMEMBER TO CHANGE TO v4!
    newSheetID = ''

    while 1:
        print('')
        print('-'*40)
        print('What would you like to do?')
        print('\t0. View A1 Notation Rules')
        print('\t1. Create new spreadsheet')
        print('\t2. View an existing spreadsheet')
        print('\t3. Append to an existing spreadsheet')
        print('\t4. Clear a portion of an existing spreadsheet')
        print('\t5. Update values of an existing spreadsheet')
        print('\t6. Exit')
        try:
            choice = int(input('Your Choice: '))
        except ValueError:
            choice = 7

        if choice == 0:
            A1_Notation()
        elif choice == 1:
            newSheetID = create(service)

        elif choice < 6:            
            print('\nEnter the ID of the spreadsheet you would like to use. Type \"NEW\" if the last new sheet created in this session has to be used.')
            spreadsheetId = input('Or enter \"TRIAL\" to use a default spreadsheet: ')
            if spreadsheetId == "TRIAL":
                spreadsheetId = '1Uk7T78T5n-HT0p9Y6ggxVF7RR06Lrn8nnZ32bsMdwhw'
                print('Note: Use sheet name \"Sheet1\" for this spreadsheet.')
            elif spreadsheetId == "NEW":
                spreadsheetId = newSheetID
                print('Note: Use sheet name \"Sheet1\" for this spreadsheet.')

            if choice == 2:
                view(service, spreadsheetId)
            elif choice == 3:
                append(service, spreadsheetId)
            elif choice == 4:
                clear(service, spreadsheetId)
            elif choice == 5:
                update(service, spreadsheetId)

        elif choice == 6:
            print('Exiting. Have a good day!')
            sys.exit()
        else:
            print('Illegal input. Kindly enter a digit between 0-6.')

def A1_Notation():
    print('\nPlease note the A1 Notation: This is a string like \"Sheet1!A1:B2\", that refers to a group of cells in the spreadsheet.')
    print('For example, valid ranges are:')
    print('\tSheet1!A1:B2 refers to the \"top-left-cell:bottom-right-cell in Sheet1\". It displays first two cells in the top two rows of Sheet1.')
    print('\tSheet1!A:A refers to all the cells in the first column of Sheet1.')
    print('\tSheet1!1:2 refers to the all the cells in the first two rows of Sheet1.')
    print('\tSheet1!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.')
    print('\tA1:B2 refers to the first two cells in the top two rows of the first visible sheet.')
    print('\tSheet1 refers to all the cells in Sheet1.')
    print('Spreadsheet API uses A1 notation for accessing spreadsheets. Fill the data keeping that in mind.')

def formRange(function):
    sheetName = input('\nEnter the name of sheet you would like to use (Enter "$1" to use first sheet): ')

    if sheetName != '$1':
        rangeName = sheetName
    else:
        rangeName = ''

    if function in ['view', 'clear', 'update']:
        cellRange = input('\nEnter the cell range as per A1 notation. \nOr enter \"all\" to %s the entire sheet (sheet name must be explicitly mentioned for this): ' %function)
        if sheetName == "$1" and cellRange.lower() == "all":
            rangeName = 'A1:Z1001'
    
    elif function == 'append':
        print('\nA sheet may have multiple tables seperated by empty rows.')
        cellRange = input('Enter any cell-address of the table, to which values have to be appended: ')

    if cellRange.lower() != "all":
        if rangeName != '':
            rangeName = rangeName + '!'
        rangeName = rangeName + cellRange
    
    print('Using data for range:', rangeName)
    return rangeName


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
        result = service.spreadsheets().create(body=data).execute()
        print('\nSpreadsheet Created! Kindly note the spreadsheet ID for further use.')
        newSheetID = result['spreadsheetId']
        print('Spreadsheet ID:', result['spreadsheetId'])
        return newSheetID
    except:
        print('Spreadsheet not created!')

def view(service, spreadsheetId):
    '''Helps in viewing data for a spreadsheet using spreadsheet ID.

    Takes input for range of data to be viewed using A1 notation- which is described in output.
    Request sent returns a json object consisting of data properties.
    The 'values' key corresponds to a list of actual spreadsheet values which is parsed and printed in terminal.
    
    For demonstration purposes, a "trial" spreadsheet is initiated and used for convenience.
    The spreadsheet ID for trial spreadsheet is hardcoded. It can be accessed at:
    https://docs.google.com/spreadsheets/d/1Uk7T78T5n-HT0p9Y6ggxVF7RR06Lrn8nnZ32bsMdwhw/edit?usp=sharing
    '''
    rangeName = formRange('view')
    majorDimension = input('\nFetch data along \"ROWS\" or \"COLUMNS\"?: ')

    try:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName, majorDimension=majorDimension).execute()
        values = result['values']
        if not values:
            print('\nNo data found.')
        else:
            print('')
            for row in values:
                curr_row = ''
                for col in row:
                    curr_row = curr_row + str(col) + ' , '
                print(curr_row)
    except:
        print('\nUnable to find requested block of data. Please check the input for sheet name and cell range.')

def append(service, spreadsheetId):
   
    rangeName = formRange('append')

    no_of_rows = int(input('\nEnter number of rows of data do you wish to append: '))
    print('Note: Data has to be entered exactly how it would be entered in a spreadsheet. Begin any formulae with \"=\" etc.')

    print('\nEnter a \"|\" separated list of data items to be appended to- ')
    for i in range(no_of_rows):
        values = []
        data_items = input('\tRow-%d: ' % (i+1))
        data = data_items.split('|')
        values.append(data)

    body = {
        'values': values
    }
    try:
        result = service.spreadsheets().values().append(spreadsheetId=spreadsheetId, range=rangeName, valueInputOption='USER_ENTERED', body=body).execute()
        print('\nData successfully appended!')
        print('Number of Rows updated: ', result['updates']['updatedRows'])
        print('Number of Columns updated: ', result['updates']['updatedColumns'])
        print('Number of Cells updated: ', result['updates']['updatedCells'])
    except:
        print('Error! Try again!')

def clear(service, spreadsheetId):
    rangeName = formRange('clear')
    
    try:
        result = service.spreadsheets().values().clear(spreadsheetId=spreadsheetId, range=rangeName, body={}).execute()
        print('Clear successful!', result['clearedRange'], 'has been cleared!')
    except:
        print('Error! Try again!')

def update(service, spreadsheetId):
    
    rangeName = formRange('update')

    no_of_rows = int(input('\nEnter number of rows of data do you wish to insert: '))
    print('Note: Data has to be entered exactly how it would be entered in a spreadsheet.\nBegin any formulae with \"=\" etc. Enter \"null\" to leave a cell unmodified')

    print('\nEnter a \"|\" separated list of data items to be appended to- ')
    for i in range(no_of_rows):
        values = []
        data_items = input('\tRow-%d: ' % (i+1))
        data = data_items.split('|')
        for i in range(data.__len__()):
            if data[i] == "null":
                 data[i] = None
            else:
                data[i].strip()
        values.append(data)
        print(data, "updated")

    body = {
      'values': values
    }
    try:
        result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range=rangeName, valueInputOption='USER_ENTERED', body=body).execute()
        print('\nData successfully updated!')
        print(result)
        print('Number of Rows updated: ', result['updatedRows'])
        print('Number of Columns updated: ', result['updatedColumns'])
        print('Number of Cells updated: ', result['updatedCells'])
    except:
        print('Error! Try Again!')

if __name__ == '__main__':
    main()