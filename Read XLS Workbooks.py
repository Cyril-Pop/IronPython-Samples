#Copyright(c) Cyril P.
#More Infos http://www.ironpython.info/index.php?title=Interacting_with_Excel
import clr

import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal

xlDirecDown = System.Enum.Parse(Excel.XlDirection, "xlDown")
xlDirecRight = System.Enum.Parse(Excel.XlDirection, "xlToRight")

input = r"C:\My Excel Files\Book1.xls"

class Lst_Xls():
	def __init__(self, path):
		ex = Excel.ApplicationClass()
		ex.Visible = False
		lst_xls = []
		workbook = ex.Workbooks.Open(path)
		ws = workbook.Worksheets[1]
		
		##get number of Rows not empty ##
		rowCountF = ws.Columns[1].End(xlDirecDown).Row
		#or
		#rowCountF = sum( x is not None for x in ws.Columns[1].Value2)
		#
		#
		##get number of Coloun not empty ##
		colCountF = ws.Rows[1].End(xlDirecRight).Column
		#or
		#colCountF = sum( x is not None for x in ws.Rows[1].Value2)
		
		for i in range(1,rowCountF+1):
			temp_lst = []
			for j in range(1,colCountF+1):
				try:
					temp_lst.append(ws.Cells[i,j].Value2.ToString())
				except:
					temp_lst.append(ws.Cells[i,j].Value2)		
			lst_xls.append(temp_lst)
		self.datas = lst_xls
		self.first_flst = [x for x in lst_xls[0]] # or lst_xls[0] 
		#Get the specify index
		self.type_fidx = self.first_flst.index("Type")
		ex.Workbooks.Close()
		ex.Quit()
    		#other proper way to make sure that you really closed and released all COM objects 
		Marshal.ReleaseComObject(workbook)
		Marshal.ReleaseComObject(ex)
		
obj_xl_lst = class Lst_Xls(input)		
