import json
from InquirerPy import prompt
from preset import Preset, Filter
import re as regex
import win32com.client
import traceback as debug
import os
from preset import preset
import sys


class MainMenu:
    
    @staticmethod
    def run() -> list:
        choices = [
            "[ Create a Sorting Preset ]",
            "[ Load a Sorting Preset ]",
            "[ Help ]",
            "[ Exit ]"
        ]
        menu = prompt(
                [
                    {
                        "type": "list",
                        "name": "usr_opt",
                        "message": "[ Would you like to create one ]",
                        "choices": choices,
                    }
                ]
            )
        choice = menu["usr_opt"]
        if choice == choices[0]:
            pass

        elif choice == choices[1]:
            pass

        elif choice == choices[2]:
            pass

        elif choice == choices[3]:
            pass



    @staticmethod
    def mainHndler():
        pass

    @staticmethod
    def creationHndler():
        # Creation = preset.CreatePreset()
        # Creation.startPresetCreation()

    @staticmethod
    def loadHndler():
        pass

    @staticmethod
    def helpHndler():
        pass

    @staticmethod
    def exitHndler():
        print("[ Exiting Program ]".center(60))
        sys.exit()


# -- Logic of Options using Presets --
class PresetHandler:
    
    '''
        Anytime the user selects create preset or load one
        we have to load there outlook client whether it be 
        to test if a folder exists or to create one. 
        This class handles the loading of the outlook client
        and the fetching of the users inbox. This class is 
        inherited by the CreatePresetHandler and LoadPresetHandler
    '''

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


class LoadPresetHandler(PresetHandler):
    ''''
        This class handles the logic when the user selects load a preset 
    
    '''
    def __init__(self) -> None:
        super().__init__()
        self.loaded_preset = None
        self.emailMove = []

    def preset_menu_options(self) -> list:
        
        '''
            We want to create a menu that displays all the json files
            in the config directory to the user and allow the user to select
            one to load in a menu
        '''

        json_files = Preset.fetch_jsons()
        if not json_files:
            print("[!] No .json Presets were found [!]")
            return None

        choices = [f"[ {file} ]" for file in json_files]
        choices.append("[ Go Back ]")
        return choices

    def run_preset_menu(self):
        
        '''
            We want to create a menu that displays all the json files
            in the PresetOptions directory and allow the user to select
            one to load
        '''

        choices = self.preset_menu_options()
        if not choices:
            print("[!] No .json Presets were found in programs config directory... [!]")
            return

        json_menu = prompt(
            [
                {
                    "type": "list",
                    "name": "usr_opt",
                    "message": "[ Select a Preset To Load ]",
                    "choices": choices,
                }
            ]
        )
        choice = json_menu["usr_opt"]
        return choice 

    def run(self):
        userJSON = self.run_preset_menu()
        if userJSON == "[ Go Back ]":
            MainMenu.run()
        else:
            self.loaded_preset = Preset.load_preset(userJSON)
            if not preset:
                print("[!] An error occurred while trying to load the user preset [!]")
                MainMenu.run()
                return 

            print("[i] Sucessfully loaded user preset [i]")
            self.create_folders_in_outlook()

    
    def create_folders_in_outlook(self):
        '''
            We want to iterate through the list of folder names
            and create them in outlook
        '''
        try:
            print("[i] Checking & Creating folders from preset file if they don't exist..  [i]")
            for folder in self.loaded_preset.folder_names:
                
                    folder = inbox.Folders.Item(folder_name)
                    # If the folder does not exist, create it
                    if folder is None:
                        self.inbox.Folders.Add(folder)

        except Exception as ex:
                print(
                    f"[!] An error occurred while trying to create the folder {folder} in Outlook [!]")
                debug.print_exc()
                return 
    
    def apply_filter(self):
        '''
            We want to iterate through the list of filters
            and move emails into the appropriate folders
        '''
        try:
            print("[i] Applying filters from preset file..  [i]")
            for filter in self.loaded_preset.folder_filters:
                folder = self.inbox.Folders.Item(filter.folder_name)
                if folder is not None:
                    for email in self.inbox.Items:
                        self.check_email(email)
                      
                    
        except Exception as ex:
            print(f"[!] An error occurred while trying to apply the filter {filter} [!]")
            debug.print_exc()
            return
    

    def check_email(self, email)

        if email.Subject in filter.subject_lines 
        or email.SenderName in filter.sender_names 
        or email.SenderEmailAddress in filter.sender_emails:
            self.emailMove.append(email)
    
    
     



class CreatePresetHandler(PresetHandler):
    def __init__(self) -> None:
        super().__init__()
        self.folderFilter = regex.compile(r'[\\/:*?"<>|]')
        self.usrFolders = []

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
            print(
                '[ Enter the name of the Folder you would like to create in Outlook ]')
            folder_name = input(
                '[?] Enter here or type "esc" to stop folder input: ').strip()
            if folder_name == 'esc':
                break
            try:
                self.inbox.Folders[folder_name]
            except Exception as ex:
                return folder_name
            print("Folde name already exists")

    def startPresetOptionCreation(self):
        '''
            We start the PresetHandler creation process
            by asking the user to enter a folder names
            until the type 'esc' and save them to folderNames
            for further processing 
        '''
        while True:
            usr_input = self.get_folder_input()
            if usr_input != 'esc':
                self.usrFolders.append(usr_input)
            else:
                break
        if not self.usrFolders:
            print(
                "[!] Cannot continue with PresetHandler creation, you must have at least one folder name")
            return

        print(self.usrFolders)

    def createFilters(self):
        '''
            We want to iterate through the list of folder names
            and ask the user for each
            1. What Subject Lines should place an email in a folder
            2. What Senders should place an email in a folder 
            3. What Email Addresses should place an email in a Folder
        '''
        for names in self.usrFolders:
            pass
        pass
