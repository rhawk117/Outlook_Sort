import json
from InquirerPy import prompt
from preset import Preset, Filter
import re as regex
import win32com.client
import traceback as debug
import os
from preset import preset
import sys
from time import sleep


class MainMenu:
    
    @staticmethod
    def run() -> list:
        
        ''''
            Runs the main menu of the program and based on the 
            users choice calls the appropriate handlers we created
            below
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
        handle_load = LoadPresetHandler()
        handle_load.run()

    @staticmethod
    def helpHndler():
        MainMenu.display_help()
        input("[i] Press Enter to Return back to the main menu... [i]".center(60))
        MainMenu.run()

    @staticmethod
    def help_message(msg: str, delay: float):
        print("[i] " + msg + " [i]", end='\n')
        sleep(delay)

    @staticmethod
    def display_help() -> None:
        MainMenu.help_message(
            "This tutorial will help you and walk you through the preset creation process for this script",
            5
        )

        MainMenu.help_message(
            """Once you start the preset creation process, you will first be asked 
        to enter at least 1 or more the folder names you want to create / sort in Outlook
        that you'd like to move emails into""",
            5
        )
        
        MainMenu.help_message(
            """NOTE: For each folder name you input, you will then have to provide at least 1 or more subject 
        lines, senders, and email addresses""",
            5
        )

        MainMenu.help_message(
            """
            The program cannot create a preset or filter for a folder if you omit any of the required fields, keep this in mind
            """, 
            5
        )

        MainMenu.help_message(
            """
            After you have inputted all the folder names and filters for the folder names and have successfully created your preset
        In the Main Menu Select 'load a preset' and the progam will load the preset you select in the menu.
            """, 
            5
        )
        
        MainMenu.help_message(
            """
            The program will then use the filters for each folder and will search your inbox for emails containing matching subject line, 
        senders, or email address and will flag any matching emails.
            """,
            5
        )

        MainMenu.help_message(
            """
            In order to avoid any mistakes, the program will ask you to confirm if you want to move the emails it has flagged for each folder
        before moving them.
            """,
            5
        )

        MainMenu.help_message(
            """
            If you would like to view this tutorial again, you can find it in the 'help' menu in the main menu
            """,
            5
        )

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
        the preset and call _create_folders_in_outlook to check if the folders
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

        choices = [f"{file}" for file in json_files]
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
            the preset and call _create_folders_in_outlook to check
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
        self._create_folders_in_outlook()
        print(f'[i] Successfully created / confirmed existence of folders in Outlook from preset file\n[i]
               applying filters [i]')
        self._apply_filters()

    def _create_folders_in_outlook(self) -> None:
        
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

    def _apply_filters(self) -> None:
        
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
            
            2). We then call _confirm_move to ask the user to confirm if they
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
        if not self._confirm_move(emails_being_moved, filter_obj.folder_name):
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
        
    def _confirm_move(self, emails_moved:list, folder:str) -> bool:
        
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
            self._hndle_view_emails(emails_moved)
            # Recursive call back to menu to confirm move after user views the emails
            return self._confirm_move(emails_moved, folder)
        
        elif choice == "[ Yes ]":
            print('[i] Moving emails... [i]')
            return True

        elif choice == "[ No ]":
            print("[i] Terminating Email Sort... [i]")
            return False

    def _hndle_view_emails(self, emails_moved:list) -> None:
        
        '''
            Displays the emails that will be moved to the user
            before the program moves them if they Select View Emails
            in the Confirm Move Menu
        '''

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
        
        - folderPattern: a pre combiled regex pattern that checks 
                         for invalid characters in a folder name
        
        - usrFolders: a list that holds the folder names that will 
                      be the keys for the nested filter objects / dictionaries 
                      inputted by the user

        - userFilters: a list that holds the filter objects / dictionaries
                          that will be the values for the nested filter objects
                         / dictionaries

    '''
    def __init__(self) -> None:
        super().__init__()
        self.folderPattern = regex.compile(r'[\\/:*?"<>|]')
        self.emailPattern = regex.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        self.known_emails: set = {emails.SenderEmailAddress for emails in self.inbox.Items}
        self.known_senders: set = {emails.SenderName for emails in self.inbox.Items}

        # Input we need from a user 
        self.userFilters: list = []
        self.usrFolders: set = []
        self.subject_lines: set = []
        self.emailAddresses: set = []
        

    def folder_name_validator(self, folder_name) -> bool:
        
        '''
            A series of checks performed to see if an inputted
            folder name can be created in Outlook true if it can, etc.
        '''

        # Check if the folder name is empty
        if not folder_name:
            print("[!] A Folder name cannot be empty. Please try again.")
            return False

        # Check for invalid characters
        if regex.search(self.folderPattern, folder_name):
            print("[!] Folder name contains invalid characters. Please try again.")
            return False
        
        # Check for duplicate folder names 
        if folder_name in self.usrFolders:
            print("[!] Folder name already exists in the list of folders in your json file. Please try again.")
            return False 

        # Check if the folder name already exists in Outlook by trying to access it, which throws if it doesn't exist
        try:
            self.inbox.Folders[folder_name]
        except Exception as ex:
            print('[i] Folder name doesn\'t exist in Outlook, saving folder input [i]')
            return True
        
        print("[!] Folder name already exists in Outlook. Please try again.")
        return False

    def email_address_validator(self, email_address: str) -> bool:
        if not regex.match(self.emailPattern, email_address):
            print(f"[!] The email address entered ({email_address}) is invalid. Please try again with a valid email address.")
            return False
        
        if email_address not in self.known_emails:
            print(f"[!] The email address you entered ({email_address}) has not sent any emails to your inbox. Please re-enter a valid email address.")
            return False
        
        if email_address in self.emailAddresses:
            print(f"[!] The email address you entered ({email_address}) already exists in the list of emails in your json file. Please try again.")
            return False
        
        return True
    
    def subject_line_validator(self, usr_subject_line: str) -> bool:
        
        '''
            We want to check if the subject line is empty or contains
            invalid characters and return True if it doesn't and False
            if it does
        '''

        if usr_subject_line.strip() == '':
            print("[i] Subject line cannot be empty. Please try again.")
            return False
        
        if not all(c.isalnum() or c.isspace() for c in usr_subject_line):
            print("[i] Subject line can only contain alphanumeric characters and spaces. Please try again.")
            return False
        
        if usr_subject_line in self.subject_lines:
            print(f"[i] Subject line already exists in the list of subject lines in your json file. Please try again.")
            return False
        
        return True

    def sender_validator(self, usr_sender: str) -> bool:
        if usr_sender.strip() == '':
            print("[i] Sender cannot be empty. Please try again.")
            return False
        if usr_sender not in self.known_senders:
            print(f"[i] The sender you entered ({usr_sender}) has not sent any emails to your inbox. Please re-enter a valid sender.")
            return False
        if usr_sender in self.senders:
            print(f"[i] The sender you entered ({usr_sender}) already exists in the list of senders in your json file. Please try again.")
            return False
        return True

    def _get_field_input(self, field_name:str, field_prompt: str, input_validator: bool):
        flag, field_input = False, ''
        while True:
            print(field_prompt)
            field_input = input(f'[?] Enter a {field_name} or "esc": ').strip()
            if field_input.lower() == 'esc':
                flag = True
                break

            elif self.input_validator(field_input):
                print(f'[i] Valid Input recieved {field_input} & it will be saved in the .json [i]')
                break

            else:
                print(f'[i] Invalid Input recieved {field_input} & it will not be saved [i]')
                continue

        return (field_input, flag)

    def get_a_fields_input(self, field_name:str, field_prompt: str, input_validator: bool) -> list:
        field_list = []
        while True:
            usr_input, esc_check = self._get_field_input(field_name, field_prompt, input_validator)
            if esc_check:
                break

            field_list.append(usr_input)

        print(f'[i] User Input for {field_name} will be saved [i]')
        print(field_list)

        return field_list

    def is_empty(self, a_list) -> bool:
        return len(a_list) == 0 or a_list is None or not a_list

    def get_folder_data(self) -> None:
        
        '''
            We start the PresetHandler creation process
            by asking the user to enter a folder names
            until the type 'esc' and save them to folderNames
            for further processing 
        '''

        # Get Folder Input 
        users_folders = self.get_a_fields_input('folder name',
                    'Enter all the folder names you\'d like this preset to create / sort from upon loading the preset or "esc" when done: ', 
                    self.folder_name_validator
        )
        if self.is_empty(users_folders):
            self.continue_menu('folder name', self.get_folder_data)
        
        self.usrFolders = users_folders  
    
    
    def continue_menu(self, field: str, current: int) -> None:
        choices = ["[ Continue ]", "[ Go Back ]"]
        yes_no_menu = prompt(
            [
                {
                    "type": "list",
                    "name": "usr_opt",
                    "message": f"[ Would you like to re-enter the {field} or go back to the Main Menu to exit Preset Creation]",
                    "choices": choices,
                }
            ]
        )
        choice = yes_no_menu["usr_opt"]
        if choice == choices[0]:
            return current
        
        elif choice == choices[1]:
            MainMenu.run()
            return -1
    
    def get_filter_data(self, current:int, current_folder:str) -> list:
        '''
        WORK IN PROGRESS
        '''
        field = []
        if current == 0:
            users_subject_lines = self.get_a_fields_input('subject line', 
                f'Enter subject lines that should move an email into the {current_folder} or "esc": ', 
                self.subject_line_validator
            )
            field = users_subject_lines

        elif current == 1:
            users_email_addresses = self.get_a_fields_input('email address', 
                f'Enter email addresses that should move an email into the {current_folder} or "esc": ', 
                self.email_address_validator 
            )
            field = users_email_addresses

        elif current == 2:
            users_senders = self.get_a_fields_input('sender', 
                f'Enter senders that should move an email into the {current_folder} or "esc": ', 
                self.email_address_validator
            )
            field = users_senders

        if not field:
            if self.continue_menu('filter', self.get_filter_data) != -1:
                self.get_filter_data(current, current_folder)
        
        return field
        
        

    def get_filters(self):
        for i, folders in enumerate(self.usrFolders, start = 0):
            self.get_filter_data(i, folders)

            






