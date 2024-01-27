import json
import os
import traceback as debug
import win32com.client
import sys
import re as regex
from preset import Preset, Filter
from InquirerPy import prompt

# -- Logic of Options using Presets --
class PresetHandler:
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
    def __init__(self) -> None:
        super().__init__()

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
