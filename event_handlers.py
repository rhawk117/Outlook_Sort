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
            hndler = LoadPresetHandler()
            hndler.run()
            
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
        and the fetching of the users inbox which is then
        inherited by the CreatePresetHandler and LoadPresetHandler
        to avoid code duplication.
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
        from the main menu, we load the outlook client and the users inbox
        and then call run_preset_menu to display the users preset options
        and allow them to select one to load, if the user selects go back
        we return to the main menu, if the user selects a preset we load
        the preset and call create_folders_in_outlook to check if the folders
        in the preset exist in outlook and create them if they don't. We then
        call apply_filter to apply the filters from the preset to the users
        inbox and then return to the main menu
    '''

    def __init__(self) -> None:
        super().__init__()
        self.loaded_preset = None

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
        if not choices or choices is None:
            print("[!] No .json Presets were found in programs config directory... [!]")
            MainMenu.run()
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
        
        '''
            Main method that runs the logic for loading a preset
            we call run_preset_menu to get the users choice of
            preset to load, if the user selects go back we return
            to the main menu, if the user selects a preset we load
            the preset and call create_folders_in_outlook to check
            if the folders in the preset exist in outlook and create
            them if they don't. We then call apply_filter to apply the
            filters from the preset to the users inbox and then return
            to the main menu
        '''
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
            self.apply_filters()

    
    def create_folders_in_outlook(self) -> None:
        
        '''
            Creates & Checks if folders exist in outlook loaded from 
            the keys of the user selected json preset file
        '''

        print("[i] Checking & Creating folders from preset file if they don't exist..  [i]")
        for folder_name in self.loaded_preset.folder_names:
            try:
                folder = self.inbox.Folders.Item(folder_name)
                if folder is None:
                    print(f"[i] Creating folder {folder_name} in Outlook [i]")
                    self.inbox.Folders.Add(folder_name)

            except Exception as ex:
                print(f"[!] An error occurred while trying to create the folder {folder_name} in Outlook [!]")
                debug.print_exc()
                return 

    def apply_filters(self) -> None:
        
        """
            Applies filters from the loaded preset file to the emails in the inbox.
            For each filter in the preset, the method finds emails that match the filter 
            criteria and asks the user to confirm if they want to move these emails to the 
            specified folder. If the user confirms, the emails are moved to the folder.
            If no emails match the filter criteria or the specified folder does not exist, 
            the method prints a message and proceeds to the next filter.
            The method prints a message for each filter to indicate whether the emails 
            were successfully moved or not.
        """

        print("[i] Applying filters from preset file...  [i]")
        for filter_obj in self.loaded_preset.folder_filters:
            try:
                # If Apply A Filter returns False at any point (i.e it failed) continue
                # to next iteration
                if not self._apply_a_filter(filter_obj):
                    continue
        
            except Exception as ex:
                print(f'[!] An error occurred while trying to apply the filter for {filter_obj.folder_name} [!]')
                debug.print_exc()
                continue

    def _apply_a_filter(self, filter_obj: 'Filter') -> bool:
        matching_emails = self._find_matches_in(filter_obj)

        if not matching_emails:
            print(f'[i] No emails were found that should be moved to {filter_obj.folder_name} [i]')
            return False
        
        # Confirm with the user that they want to move the emails
        if self.confirm_move(matching_emails, filter_obj.folder_name):
            folder = self.inbox.Folders.Item(filter_obj.folder_name)

            # Double Check folder exists 
            if folder is not None:
                for email in matching_emails:
                    email.Move(folder)
                print(f'[i] Sucessfully moved emails to {filter_obj.folder_name} [i]')
                return True

            else:
                print(f'[!] Could not find a folder named "{filter_obj.folder_name}" [!]')
                return False

        else:
            print(f'[i] Proceeding to next folder [i]')
            return False

    def _find_matches_in(self, filter_obj: 'Filter') -> list:
        """
            Finds emails in the inbox that match the criteria specified in the given filter object.

            The method checks each email in the inbox against the subject lines, senders, and email addresses specified in the filter.
            If an email matches any of the criteria, it's added to the list of matching emails.

            Args:
                filter_obj (Filter): The filter object containing the criteria to match against.

            Returns:
                list: A list of emails that match the filter criteria.
        """
        
        matching_emails = []
        for subject_line in filter_obj.subject_lines:
            matching_emails.extend(self.inbox.Items.Restrict(f"[Subject] = '{subject_line}'"))

        for sender in filter_obj.senders:
            matching_emails.extend(self.inbox.Items.Restrict(f"[SenderName] = '{sender}'"))

        for email_address in filter_obj.emails:
            matching_emails.extend(self.inbox.Items.Restrict(f"[SenderEmailAddress] = '{email_address}'"))

        return matching_emails

    def confirm_move(self, emails_moved:list, folder:str) -> bool:
        print(f"[>>] Total emails to move: {len(emails_moved)}\n")
        input("[i] Press enter to continue, you will be asked if you'd like to move the emails. [i]")
        yes_no_menu = prompt(
            [
                {
                    "type": "list",
                    "name": "usr_opt",
                    "message": f"[ Would you like to move the emails found into {folder}, this action cannot be undone ]",
                    "choices": ["[ Yes ]", "[ No ]"],
                }
            ]
        )

        choice = yes_no_menu["usr_opt"]
        if choice == "[ Yes ]":
            print('[i] Moving emails... [i]')
            return True

        elif choice == "[ No ]":
            print("[i] Terminating Email Sort... [i]")
            return False




    
    
     



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
