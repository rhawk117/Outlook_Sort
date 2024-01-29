# -- Preset & Filter Objects -- #
import json
import os
import traceback as debug
from dataclasses import dataclass


@dataclass
class Filter:
    '''
        Class represents each of the nested dictionaries
        in the preset .json file
    '''
    subject_lines: list
    emails: list
    senders: list
    folder_name: str

    def to_dict(self) -> dict:
        return {
            'subject_lines': self.subject_lines,
            'emails': self.emails,
            'senders': self.senders,
            'folder_name': self.folder_name
        }

    def is_valid(self) -> bool:
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

            return Preset(
                list(preset.keys()),
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

