#Copyright(c) Cyril P.
#More Infos http://www.ironpython.info/index.php?title=Interacting_with_Excel
import clr
import sys
import System
from System import Array
from System.Collections.Generic import *

clr.AddReference('System.Windows.Forms')
import System.Windows.Forms
from System.Windows.Forms import SaveFileDialog, DialogResult

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal

xlTypePDF = Excel.XlFixedFormatType.xlTypePDF


class Xls_Utils():
	def __init__(self, path ,nameSheet = 1):
		self._path = path
		self._nameSheet = nameSheet
		
	def ExportPdf(self):
		ex = Excel.ApplicationClass()
		ex.Visible = False
		lst_xls = []
		workbook = ex.Workbooks.Open(Filename=self._path, ReadOnly=True)
		ws = workbook.Worksheets[self._nameSheet]
		ws.Activate()
		pages = ws.PageSetup.Pages.Count
		print(pages)
		saveXlsxFileDialog = SaveFileDialog()
		saveXlsxFileDialog.Title = 'Name of PDF'
		saveXlsxFileDialog.FileName = "{}_{}".format(System.IO.Path.GetFileNameWithoutExtension(self._path), str(self._nameSheet))
		saveXlsxFileDialog.DefaultExt = "pdf"
		saveXlsxFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*"
		saveXlsxFileDialog.RestoreDirectory = True
		saveXlsxFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(self._path)
		resultSaveAs = saveXlsxFileDialog.ShowDialog()
		if resultSaveAs == DialogResult.OK:
			fullpath = saveXlsxFileDialog.FileName 
			workbook.ExportAsFixedFormat(Type=xlTypePDF, FileName=fullpath, From=1, To=pages)
		ex.Workbooks.Close()
		ex.Quit()
		#other proper way to make sure that you really closed and released all COM objects 
		Marshal.ReleaseComObject(workbook)
		Marshal.ReleaseComObject(ex)


input = IN[0] # excel full path
xls_sheetName = IN[1] # name of sheet
obj_xls = Xls_Utils(input, xls_sheetName)
obj_xls.ExportPdf()
