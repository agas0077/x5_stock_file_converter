from tkinter import filedialog, messagebox, Entry
from ntpath import basename
from os import path


class Dialog:
    def __init__(self):
        self.fileNames = []
        self.pathTuple = ()
        self.formatFileName = r""

    def callDialog(self):
        """Calls open file dialog, possible to choose only '.xlsx .xls .xlsm .xlsb'"""
        self.pathTuple = filedialog.askopenfilenames(filetypes=[("Excel files", ".xlsx .xls .xlsm .xlsb")])
        self.fileNames = [basename(path.abspath(name)) for name in self.pathTuple]
    
    def callDialogOneFile(self):
        self.formatFileName = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls .xlsm .xlsb")])

    def getPaths(self):
        """Returns tuple of paths stored at class instance"""
        return self.pathTuple

    def getFormatFile(self):
        return self.formatFileName

    


    

