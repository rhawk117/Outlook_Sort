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
        
        ''''
        
        '''

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
        pass

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

        # No json files were found in the config directory
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
        # Recieve the users choice of the preset to load
        userJSON = self.run_preset_menu()
        if userJSON == "[ Go Back ]":
            MainMenu.run()
            return
        
        self.loaded_preset = Preset.load_preset(userJSON)
        if not preset:
            print("[!] An error occurred while trying to load the user preset [!]")
            MainMenu.run()
            return 

        print("[i] Sucessfully loaded user preset [i]")
        self.create_folders_in_outlook()
        print(f'[i] Successfully created / confirmed existence of folders in Outlook from preset file\n[i]
               applying filters [i]')
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
                    print(f"[i] Creating folder {folder_name} in Outlook since it didn't exist [i]")
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
        
        '''
            Performs the series of checks & actions performed 
            when applying a singular filter to the users inbox
            refactored into a method to improve code readability

            1). We first try to find emails that match the filter criteria
                if no emails are found we print a message and return False
            
            2). We then call confirm_move to ask the user to confirm if they
                want to move these emails to the specified folder. 
                    -If they select No we return False and continue
                     to the next filter
                        
            
            3). If the user confirms we should move the emails the program found
                double check if the folder exists in outlook which it always should
                unless the user deleted it manually. 
                    -If the folder doesn't exist we return False and continue to 
                    the next filter
            
            4). If the folder exists we call _move_emails to move the emails to the folder
        '''

        emails_being_moved = self._find_matches_in(filter_obj)

        # No Emails were found that matched the filter criteria
        if not emails_being_moved:
            print(f'[i] No emails were found that should be moved to {filter_obj.folder_name} [i]')
            return False
        
        # Confirm with the user that they want to move the emails
        if not self.confirm_move(emails_being_moved, filter_obj.folder_name):
            return False 
        
        folder = self.inbox.Folders.Item(filter_obj.folder_name)

        # Double Check folder exists before email is moved 
        if folder is None:
            print(
                f'[!] Could not find a folder named "{filter_obj.folder_name}" [!]'
            )
            return False
        
        self._move_emails(emails_being_moved, folder)
        print(f'[i] Sucessfully moved {len(emails_being_moved)} emails to {filter_obj.folder_name} [i]')
        return True    

        

    def _move_emails(self, emails_being_moved: list, folder: str) -> None:
        
        '''
            Moves the given list of emails into the specified folder.
        '''

        for email in emails_being_moved:
            try:
                email.Move(folder)
            except Exception as ex:
                print(f'[!] An error occurred while trying to move the email {email.Subject} to {folder} [!]')
                debug.print_exc()
                continue

    def _find_matches_in(self, filter_obj: 'Filter') -> list:
        
        """
            Finds emails in the inbox that match the criteria specified in the given filter object.

            The method checks each email in the inbox against the subject lines, senders, and email 
            addresses specified in the filter.
            
            - If an email matches any of the criteria, it's added to the list of matching emails.

            Args:
                filter_obj (Filter): The filter object containing the criteria to match against.

            Returns:
                list: A list of emails that match the filter criteria.
        """
        
        emails_being_moved = []
        # Check each email in the inbox against the filters list of subject lines 
        for subject_line in filter_obj.subject_lines:
            emails_being_moved.extend(
                self.inbox.Items.Restrict(f"[Subject] = '{subject_line}'")            
            )

        # Check each email in the inbox against the filters list of senders
        for sender in filter_obj.senders:
            emails_being_moved.extend(
                self.inbox.Items.Restrict(f"[SenderName] = '{sender}'")
            )

        # Check each email in the inbox against the filters list of email addresses
        for email_address in filter_obj.emails:
            emails_being_moved.extend(
                self.inbox.Items.Restrict(f"[SenderEmailAddress] = '{email_address}'")                
            )

        return emails_being_moved

    def _display_an_email(self, email) -> None:
        print(f"[>>] Email Subject: {email.Subject}")
        print(f"[>>] Email Sender: {email.SenderName}")
        print(f"[>>] Email Sender Email Address: {email.SenderEmailAddress}\n")
        
        

    def confirm_move(self, emails_moved:list, folder:str) -> bool:
        
        '''
            Upon finding emails that match the filter criteria, this method asks 
            the user to confirm if they want to move these emails to the specified folder.

            This ensures that if the program makes an error the user can cancel the moving action
            before it happens.
        '''

        menu_items = ["[ View Emails Being Moved ]", "[ Yes ]", "[ No ]"]
        menu_prompt = f"[ Would you like to move the {len(emails_moved)} emails" + f" found into {folder} ]",
        warning = "\n[!] WARNING this action cannot be undone [!]"
        yes_no_menu = prompt(
            [
                {
                    "type": "list",
                    "name": "usr_opt",
                    "message": menu_prompt + warning,
                    "choices": menu_items,
                }
            ]
        )

        choice = yes_no_menu["usr_opt"]
        
        # User wants to view the emails being moved
        if choice == "[ View Emails Being Moved ]":
            self.hndle_view_emails(emails_moved)
            # Recursive call back to menu to confirm move after user views the emails
            return self.confirm_move(emails_moved, folder)
        
        elif choice == "[ Yes ]":
            print('[i] Moving emails... [i]')
            return True

        elif choice == "[ No ]":
            print("[i] Terminating Email Sort... [i]")
            return False

    def hndle_view_emails(self, emails_moved:list) -> None:
        for email in emails_moved:
            self._display_an_email(email)

        print("*"*60)
        input(
            "\n" + "[i] Press Enter to Return back to the confirmation menu... [i]".center(60) + "\n"
        )
        print("*"*60)



    
    
class CreatePresetHandler(PresetHandler):
    '''
        This class handles the logic when the user selects create a preset,

        We have the following attributes for our class
        
        - folderPattern: a pre combiled regex pattern that checks for invalid characters in a folder name
        
        - usrFolders: a list that holds the folder names that will be the keys for the nested filter objects / dictionaries 
                      inputted by the user

        - userFilters: a list that holds the filter objects / dictionaries that will be the values for the nested filter objects
                         / dictionaries

    '''
    def __init__(self) -> None:
        super().__init__()
        self.folderPattern = regex.compile(r'[\\/:*?"<>|]')
        self.usrFolders: list = []
        self.userFilters: list = []

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
            print("[!] A Folder name cannot be empty. Please try again.")
            return False

        # Check for invalid characters
        if regex.search(self.folderPattern, folder_name):
            print("[!] Folder name contains invalid characters. Please try again.")
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
        while True:
            print(
                '[ Enter the name of the Folder you would like to create in Outlook ]')
            folder_name = input(
                '[?] Enter here or type "esc" to stop folder input: '
                ).strip()
                
            if folder_name == 'esc':
                break

            # We recieved invalid folder name input from the user
            if not self.folder_name_checks(folder_name):
                continue

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
