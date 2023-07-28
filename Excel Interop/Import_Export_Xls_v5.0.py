# update 15/04/2023 for ipy3
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

import System
clr.AddReference("System.Numerics")
from System import Array
from System.Collections.Generic import *

clr.AddReference('System.Drawing')
import System.Drawing
from System.Drawing import *

clr.AddReference('System.Data')
from System.Data import *


try:
    clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
    from Microsoft.Office.Interop import Excel
except Exception as ex:
    try:
        clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
        from Microsoft.Office.Interop import Excel
    except Exception as ex:
        Excel = typelib.Excel  
	
"""
try:
	clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' )
except:
	clr.AddReference('Microsoft.Office.Interop.Excel')
"""
# from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal

# method from Interrop dll GAC
try:
	xlDirecDown = System.Enum.Parse(Excel.XlDirection, "xlDown")
	xlDirecRight = System.Enum.Parse(Excel.XlDirection, "xlToRight")
	xlDirecUp = System.Enum.Parse(Excel.XlDirection, "xlUp")
	xlDirecLeft = System.Enum.Parse(Excel.XlDirection, "xlToLeft")
# method from Microsoft.Scripting.ComInterop.ComTypeLibInfo
except:
	xlDirecDown = Excel.XlDirection.xlDown
	xlDirecRight = Excel.XlDirection.xlToRight
	xlDirecUp = Excel.XlDirection.xlUp
	xlDirecLeft = Excel.XlDirection.xlToLeft

class ExcelUtils():	
	@staticmethod
	def ConvertDataTableToArray(dataTable):
		"""
		This function converts a DataTable object to an Array.
		:param dataTable: DataTable object to be converted
		:type dataTable: DataTable
		:return: Array of data from DataTable object
		:rtype: Array
		"""
		arrayColumns = [c.ColumnName for c in dataTable.Columns]
		arrayRows = [[None if isinstance(j, System.DBNull) else j for j in row.ItemArray ] for row in dataTable.Rows]
		arrayRows.insert(0, arrayColumns)
		return arrayRows
		
	@staticmethod
	def ConvertArrayToDataTable(lstdata):
		dataTable = DataTable()
		# create columns
		for idx, item in enumerate(lstdata[0]):
			dataTable.Columns.Add(item)
		# add rows
		for sublst_values in lstdata[1:]:
			a = Array.CreateInstance(System.Object, len(sublst_values))
			for i, val in enumerate(sublst_values):
				a[i] = val
			dataTable.Rows.Add(a)
		return dataTable
		
		
	@staticmethod
	def ExportXls(filepath, array_data, first_line_asHeader = True):
		"""
		This function exports data from an Array to an Excel workbook.
		:param filepath: full path of out excel file
		:type filepath: str
		
		:param array_data: Array of data to be exported
		:type array_data: nested List or Array
		:return: None
		:rtype: None
		"""
		ex = Excel.ApplicationClass()
		#
		ex.Visible = True
		ex.DisplayAlerts = False
		workbook = ex.Workbooks.Add()
		workbook.SaveAs(filepath)
		ws = workbook.Worksheets[1]	
		nbr_row = len(array_data)
		nbr_colum = len(array_data[0])
		xlrange  = ws.Range[ws.Cells(1, 1), ws.Cells(nbr_row, nbr_colum)]
		a = Array.CreateInstance(object, nbr_row, nbr_colum)
		for indexR, row in enumerate(array_data):
			for indexC , value in  enumerate(row):
				a[indexR,indexC] = System.Int32(value) if isinstance(value, System.Numerics.BigInteger) else value
				
		#copy Array in range			
		xlrange.Value2 = a		
		used_range = ws.UsedRange	
		for column in used_range.Columns:
			column.AutoFit()
		# apply style
		missing = System.Type.Missing
		try:
			if first_line_asHeader:
				new_table = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, xlrange, missing, Excel.XlYesNoGuess.xlYes, missing)
			else:
				new_table = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, xlrange, missing, Excel.XlYesNoGuess.xlNo, missing)
				#new_table = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, xlrange, missing, Excel.XlYesNoGuess.xlGuess, missing)
			new_table.Name = "WFTableStyle"
			new_table.TableStyle = "TableStyleMedium6"
		except:pass
			
	@staticmethod
	def ImportXls(filepath, lst_ColumnName = [], return_dataTable = False, sheetName = None):
		"""
		This function imports an excel file into a DataTable or a list of lists
		:param filepath: full path of out excel file
		:type filepath: str
		
		:param lst_ColumnName: list of column names to be imported
		:param return_dataTable: boolean indicating whether to return a DataTable or a list of lists
		:param sheetName: name of the sheet to be imported
		
		:return: DataTable or a list of lists
		"""
		workbook = None
		ws = None
		datas = None
		error = None
		#
		ex = Excel.ApplicationClass()
		#
		ex.Visible = False
		lst_xls = []
		try:
			workbook = ex.Workbooks.Open(filepath)
			if sheetName is not None:
				try:
					ws = workbook.Worksheets[sheetName]
				except Exception as ex:
					print(ex)
					ws = workbook.Worksheets[1]
			else:
				ws = workbook.Worksheets[1]
			################################
			##get number of Rows not empty ##
			#################################
			# get Rows Count method 1
			# rowCountF = max(ws.Range(i).End(xlDirecUp).Row for i in ["A65536", "B65536", "C65536", "D65536", "E65536", "F65536", "G65536", "H65536"])
			# get Rows Count method 2
			# rowCountF = ws.Columns[1].End(xlDirecDown).Row
			# get Rows Count method 3
			#rowCountF=ws.UsedRange.Rows.Count
			####################################
			## get number of Columns not empty ##
			#####################################
			# get Columns Count method 1
			# colCountF = max(ws.Range(i).End(xlDirecLeft).Column for i in ["ZZ1", "ZZ2", "ZZ3", "ZZ4", "ZZ5", "ZZ6", "ZZ7", "ZZ8", "ZZ9"])
			# get Columns Count method 2
			# colCountF = ws.Rows[1].End(xlDirecRight).Column
			# get Columns Count method 3
			#colCountF=ws.UsedRange.Columns.Count
			#####################################
			### other  method 2 maybe the best ###
			######################################
			last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, System.Type.Missing)
			usedrange = ws.Range["A1", last]
			rowCountF = last.Row
			colCountF = last.Column
			#print((rowCountF, colCountF))
			#
			fullrange = ws.Range[ws.Cells(1, 1), ws.Cells(rowCountF, colCountF)]
			fullvalue = list(fullrange.Value2)
			#split list into sublist with number of colum
			n = colCountF					
			datas = list(fullvalue[i:i+n] for i in range(0, len(fullvalue ), n))
			# 
			# convert to DataTable
			dt = ExcelUtils.ConvertArrayToDataTable(datas)
			# if lst_ColumnName is not empty remove other column by name
			if len(lst_ColumnName) > 0:
				# convert to DataView
				dataView = dt.DefaultView
				# re-convert to DataTable  with selection of Columns
				dt = dataView.ToTable(False, System.Array[System.String](lst_ColumnName))
			# overide data
			if return_dataTable:
				datas = dt
			else:
				datas = ExcelUtils.ConvertDataTableToArray(dt)
		except Exception as excp:
			import traceback
			error = traceback.format_exc()
		#
		# close excel properly
		if ex is not None:
			ex.Workbooks.Close()
			ex.Quit()
		#other proper way to make sure that you really closed and released all COM objects 
		if workbook is not None:
			Marshal.ReleaseComObject(workbook)
		if ex is not None:
			Marshal.ReleaseComObject(ex)
		workbook = None
		ex = None
		if error is not None:
			return error
		else:
			return datas
		
fileapth = IN[0]
#datas = ExcelUtils.ImportXls(fileapth, lst_ColumnName = ["Employee ID","Full Name","Job Title"], return_dataTable = False)
datas = ExcelUtils.ImportXls(fileapth, return_dataTable = False, lst_ColumnName = ["Element type Name", "Layer Name"], sheetName = "Non parameters")
# if return_dataTable = True
if isinstance(datas, DataTable):
	OUT = [c.ColumnName  for c in datas.Columns]
else:
	# new_file_path = fileapth.replace(".xlsx", "_v2.xlsx")
	# export_datas = ExcelUtils.ExportXls(new_file_path, datas, first_line_asHeader = False)
	OUT = datas
