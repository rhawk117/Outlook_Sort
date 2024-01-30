import os 
import sys
import logging as log
from preset import Preset
from InquirerPy import prompt
from event_handlers import MainMenu




class App:
    def __init__(self) -> None:
        self._startUp()
        self.uponStart()

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
        MainMenu.run()

            


    

 



