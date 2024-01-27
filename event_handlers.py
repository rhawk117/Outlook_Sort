import preset
import sys


class Handler:
    @staticmethod
    def mainHndler():
        pass

    @staticmethod
    def creationHndler():
        Creation = preset.CreatePreset()
        Creation.startPresetCreation()

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