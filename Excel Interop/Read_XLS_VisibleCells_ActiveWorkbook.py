import sys
import clr
import System

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal

specsVisu = System.Enum.Parse(Excel.XlCellType, "xlCellTypeVisible")

class Lst_Xls():
	def __init__(self):
		try:   
			ex = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
		except: 
			ex = Excel.ApplicationClass()   
		ex.Visible = True
		self.app = ex 
		self.error = None
		self.data = []  
		self.workbook = ex.ActiveWorkbook
		self.ws = ex.ActiveSheet
		
	def ReadVisibleCells(self):
		lst_xls = []
		plagefiltrevisible = self.ws.UsedRange.SpecialCells(specsVisu).Rows
		for row in plagefiltrevisible:
			lst_xls.append(row.Value2)
		return lst_xls

objxls = Lst_Xls()	

OUT = objxls.ReadVisibleCells()
