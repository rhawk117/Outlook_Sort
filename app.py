import os 
import sys
import logging as log
from preset import Preset
import win32com.client
from InquirerPy import prompt
import preset 
from event_handlers import Handler




class App:
    def __init__(self) -> None:
        self._startUp()

    def _startUp(self):
        
        '''
            initialize app properties / class members for 
            ez re-use
        '''

        self.config = 'config'
        self.progPath = os.path.abspath(os.path.dirname(sys.argv[0]))
        self.configPath = os.path.join(self.progPath, self.config)
        

    def uponStart(self):
        
        '''
            Change to script directory & check for json
            presets 
        '''

        os.chdir(self.progPath)
        if not os.path.exists(self.config):
            print('[i] Config directory not found creating it at program path...')
            os.mkdir(self.config)

        log.basicConfig(filename=f'{self.config}\\program.log', level=log.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s')
        log.info('Finished program configurations')
    
    def MainMenu(self):
        choices = [
            "[ Create a Sorting Preset ]",
            "[ Load a Sorting Preset ]",
            "[ Help ]",
            "[ Exit ]"
        ]

        yes_no_menu = prompt(
            [
                {
                    "type": "list",
                    "name": "usr_opt",
                    "message": "[ Welcome Select a Menu Option ]",
                    "choices": choices,
                }
            ]
        )
        choice = yes_no_menu["usr_opt"]
        if choice == choices[0]:
            Handler.creationHndler()

        elif choice == choice[1]:
            Handler.loadHndler()

        elif choice == choice[2]:
            Handler.helpHndler()

        elif choice == choice[3]:
            Handler.exitHndler()

    def run(self):
        self.MainMenu()
            


    

 



