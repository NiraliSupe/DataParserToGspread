import gspread, json
from collections import OrderedDict
from oauth2client.service_account import ServiceAccountCredentials

class gdataImport(object):

    def __init__(self, credentialFile, scope):
        # use creds to create a client to interact with the Google Drive API
        self.creds       = ServiceAccountCredentials.from_json_keyfile_name(credentialFile, scope)
        self.client      = gspread.authorize(self.creds)

    def createSpreadsheet(self):
        try:
        # open spreadsheet if exists else create
           spreadsheet  = self.client.open("gdata-import")
        except:
           spreadsheet  = self.client.create("gdata-import")        
           #To be able to access newly created spreadsheet, share it with email.
           spreadsheet.share('nps287@nyu.edu', perm_type='user', role='writer')

        try:
           # By title
           worksheet = spreadsheet.worksheet("data-parse")
           spreadsheet.del_worksheet(worksheet)
        except:
           pass
        finally:
           worksheet = spreadsheet.add_worksheet(title="data-parse", rows="100", cols="10")
        return worksheet
 
    def readJsonFile(self):
        print (" ** Retrieving data...")
        data_file = open('data.json', 'r')
        data = json.load(data_file, object_pairs_hook=OrderedDict)
        return data
        
    def writeToSpreadSheet(self): 
        worksheet = self.createSpreadsheet()
        # Retrive the json data
        data      = self.readJsonFile()
        print (" ** Writting data in spreadsheet")
        index     = 1
        header    = ['Name', 'view_grades', 'change_grades', 'add_grades', 'delete_grades', 'view_classes', 'change_classes', 'add_classes', 'delete_classes']
        # Write header
        worksheet.insert_row(header, index)
        # Write data to worksheet
        for name, permissions in data.items():
           row = [name]+ [0]*(len(header)-1)
           for perm in permissions:
               idx      = header.index(perm) 
               row[idx] = 1
           index = index + 1
           worksheet.insert_row(row, index)

if __name__ == "__main__":
    scope          = ['https://spreadsheets.google.com/feeds',  'https://www.googleapis.com/auth/drive']
    credentialFile = 'gdata_import-08b7b7b3dc70.json'
    gd             = gdataImport(credentialFile, scope)
    gd.writeToSpreadSheet()
    print (" ** Goodbye")
