import json 
import os 
import traceback as debug
import win32com.client
import sys
import re as regex



class Preset:
    def __init__(self) -> None:
        self.load_outlook_client()
        self.inbox = self.Outlook.GetDefaultFolder(6)

    def load_outlook_client(self):
        try: 
            self.Outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        except Exception:
            print('[!] An error occured while trying to open your Outlook client ensure that Outlook is installed before proceeding [!]')
            sys.exit()
        print("[i] Sucessfully loaded Outlook client [i]")

    def fetch_jsons():
        
        '''
            retrieves all json files from presets dir
            and returns them
        '''

        # Lambda expression to fetch previously create json presets
        json_files = list(filter
            (
                lambda file: file.endswith('.json'), os.listdir('config')
            )
        )
        return json_files
    
    def open_preset(file_name:str) -> dict:
        '''
            opens the json file in param in the config dir
            and returns it as a dictionary
        '''
        
        preset = {}
        try:
            with open(file_name, mode = 'r') as file:
                preset = json.load(file)
        except Exception:
            print(f'[!] An error occured while trying to open {file_name} details on line below[!]')
            debug.print_exc()

        return preset
        
    

class LoadPreset(Preset):
    def __init__(self) -> None:
        super().__init__()

    

class CreatePreset(Preset):
    def __init__(self) -> None:
        super().__init__()
        self.folderFilter = regex.compile(r'[\\/:*?"<>|]')
        self.folderNames = []

    def folder_name_checks(self, folder_name) -> bool:
        
        '''
            A series of checks performed to see if an inputted
            folder name can be created in Outlook true if it can, etc.
        '''
        
        if folder_name == '':
            return False
        if folder_name == 'esc':
            return True
        
        # Check if the folder name is empty
        if not folder_name:
            print("Folder name cannot be empty. Please try again.")
            return False

        # Check for invalid characters
        if regex.search(self.folderFilter, folder_name):
            print("Folder name contains invalid characters. Please try again.")
            return False
    
        return True

    def get_folder_input(self):
        
        '''
            While loop that uses previous function above
            and a try block to determine whether or not
            an inputted folder name already exists in outlook
            returns the string inputted upon recieving valid 
            input from the user
        '''

        folder_name = ''
        while not self.folder_name_checks(folder_name):
            print('[ Enter the name of the Folder you would like to create in Outlook ]')
            folder_name = input('[?] Enter here or type "esc" to stop folder input: ').strip()
            if folder_name == 'esc':
                break
            try:
                self.inbox.Folders[folder_name]
            except Exception as ex:
                return folder_name
            print("Folde name already exists")
            
    
    def startPresetCreation(self):
        
        '''
            We start the preset creation process
            by asking the user to enter a folder names
            until the type 'esc' and save them to folderNames
            for further processing 
        '''
        while True:
            usr_input = self.get_folder_input()
            if usr_input != 'esc':
                self.folderNames.append(usr_input)
            else:
                break
        if not self.folderNames:
            print("[!] Cannot continue with preset creation, you must have at least one folder name")
            return
        
        print(self.folderNames)
    
    def createFilters(self):
        '''
            We want to iterate through the list of folder names
            and ask the user for each
            1. What Subject Lines should place an email in a folder
            2. What Senders should place an email in a folder 
            3. What Email Addresses should place an email in a Folder
        '''
        for names in self.folderNames:
            pass
        pass
        
        



    

    
    

    


