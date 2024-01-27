import json 
import os 
import traceback as debug
import win32com.client
import sys
import re as regex

# -- Preset & Filter Objects -- #

import json
import os
import traceback as debug
import win32com.client
import sys
import re as regex
from dataclasses import dataclass


@dataclass
class Filter:
    subject_lines: list
    emails: list
    senders: list
    folder_name: str

    def to_dict(self):
        return {
            'subject_lines': self.subject_lines,
            'emails': self.emails,
            'senders': self.senders
        }

    def is_valid(self):
        return self.subject_lines or self.emails or self.senders


class Preset:
    def __init__(self, folders: list, filters: list, file_name: str):
        self.as_json = {}
        self._set_preset(folders, filters, file_name)

    def _set_preset(self, folders, filters, file_name):
        '''
            sets the preset properties used in class constructor
        '''

        if len(folders) != len(filters):
            print('[!] The number of folders and filters do not match [!]')
            return

        self.folder_names = folders
        self.preset_filters = filters
        self.file_name = file_name
        # Sets the as_json property to a dictionary
        self._construct_preset()

    def _construct_preset(self):
        # We can't use zip() aren't the same length
        if len(self.folder_names) != len(self.preset_filters):
            print('[!] The number of folders and filters do not match [!]')
            return

        for folder_name, filters in zip(self.folder_names, self.preset_filters):
            if filters.is_valid():
                self.as_json[folder_name] = filters.to_dict()

    def save_preset(self):
          """
                Save the user preset as a JSON file.

                This method checks if the preset can be saved as JSON and if the 'config' directory exists.
                If the conditions are met, it saves the user preset as a JSON file in the 'config' directory.

                Returns:
                    None

                Raises:
                    Exception: If an error occurs while trying to save the user preset.
            """

           if not self.as_json or not os.path.exists('config'):
                print(
                    '[!] Cannot save user preset due to an issue constructing the JSON [!]')
                return

            try:
                with open(f'config\\{self.file_name}', mode='w') as file:
                    json.dump(self.as_json, file, indent=4)
                print('[i] Sucessfully saved user preset [i]')

            except Exception as ex:
                print(
                    '[!] An error occurred while trying to save the user preset [!]')
                debug.print_exc()
                return None

    @staticmethod
    def load_preset(preset_name:str) -> 'Preset':
        """
            Load a preset from a JSON file.

            This method reads a JSON file whose name is given by `preset_name`, 
            and constructs a `Preset` object from the data in the file. The keys 
            in the JSON file become the folder names in the `Preset`, and the values 
            become the filters.

            If an error occurs while trying to load the preset, an error message 
            is printed and the method returns `None`.

            Parameters
            ----------
            preset_name : str
                The name of the JSON file to load the preset from.

            Returns
            -------
            Preset
                The `Preset` object constructed from the data in the JSON file, 
                or `None` if an error occurred.
        """

        try:
            preset = {}
            with open(f'config\\{preset_name}', mode='r') as file:
                preset = json.load(file)

            return Preset(list(preset.keys()),
                          Preset._fetch_preset_filters(preset),
                          preset_name
                        )

        except Exception as ex:
            print('[!] An error occurred while trying to load the user preset [!]')
            debug.print_exc()
            return None

    @staticmethod
    def _fetch_preset_filters(loaded_json: dict):
        # unpack all the nested dictionaries into filter objects so we can use them
        return [Filter(**filter_data) for filter_data in loaded_json.values()]
    
    @staticmethod
    def fetch_jsons():
    
    '''
        retrieves all json files from PresetOptions dir
        and returns them
    '''

    # Lambda expression to fetch previously create json PresetOptions
    json_files = list(filter
        (
            lambda file: file.endswith('.json'), os.listdir('config')
        )
    )
    return json_files
    


# -- Example of a JSON Preset -- #
# {
#     "Example Folder 1": {
#         "Subject Lines": ["3320", "TCP/IP"],
#         "Emails": ["example@outlook.com", "example@outlook.com"],
#         "Senders": ["John Doe", "Jane Doe"]
#     },
#     "Example Folder 2": {
#         "Subject Lines": ["3320", "TCP/IP"],
#         "Emails": ["example@outlook.com", "example@outlook.com"],
#         "Senders": ["John Doe", "Jane Doe"]
#     }

# }



# -- Logic of Options using Presets --
class PresetOption:
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
    


class LoadPresetOption(PresetOption):
    def __init__(self) -> None:
        super().__init__()

    def preset_menu_options(self):
        '''
            We want to create a menu that displays all the json files
            in the PresetOptions directory and allow the user to select
            one to load
        '''
        json_files = Preset.fetch_jsons()
        if not json_files:
            print("[!] No PresetOptions found [!]")
            return
        choices = []
        for file in json_files:
            choices.append(f"[ {file} ]")
        choices.append("[ Go Back ]")
        return choices
    
    def preset_menu(self):
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

        return json_menu
    

class CreatePresetOption(PresetOption):
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
            
    
    def startPresetOptionCreation(self):
        
        '''
            We start the PresetOption creation process
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
            print("[!] Cannot continue with PresetOption creation, you must have at least one folder name")
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
        
        



    

    
    

    


