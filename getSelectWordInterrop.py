import sys
import clr
import System

clr.AddReference("Microsoft.Office.Interop.Word")
import Microsoft.Office.Interop.Word as Word

clr.AddReference("System.Runtime.InteropServices")
import System.Runtime.InteropServices


class WordUtils():
    def __init__(self):
        try:
            self.wapp = Word.ApplicationClass()
            self.document = self.wapp.ActiveDocument
            self.docName = self.document.FullName
        except:
            self.interrop = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")  
            self.wapp = self.interrop.Application
            self.document = self.wapp.ActiveDocument
            self.docName = self.document.FullName
    
    @property   
    def getSelection(self):
        return self.wapp.Selection.Text
         
        
objWord = WordUtils()   


OUT = objWord.getSelection
