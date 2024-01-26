import json 
import os 
import traceback as debug

class Preset:
    def __init__(self) -> None:
        self.avlblePresets: list = self.fetch_jsons()
        self.havePresets: bool = len(self.avlblePresets) > 0

    def fetch_jsons():
        '''
            retrieves all json files from presets dir
            and returns them
        '''
        # Lambda expression to fetch previously create json presets
        json_files = list(filter(lambda file: file.endswith(
            '.json'), os.listdir('config')))
        return json_files
    
    def open_preset(file_name:str):
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
    pass

class CreatePreset(Preset):
    pass
    


