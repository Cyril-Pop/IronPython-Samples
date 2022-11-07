# Charger les bibliothèques DesignScript et Standard Python
import sys
import clr
import System
#import net library
from System import Array
from System.Collections.Generic import List, IList, Dictionary
clr.AddReference('System.Data')
clr.AddReference('System.Data.DataSetExtensions')
from System.Data import *
clr.ImportExtensions(System.Data.DataTableExtensions)

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
import Autodesk.DesignScript.Geometry as DS

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.DB.Mechanical import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
uiapp=DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
sdk_number = int(app.VersionNumber)

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import BindingSource 
clr.AddReference("IronPython.Wpf")
import wpf
from System import Windows
from System.Windows.Controls import *
from System.IO import StringReader
clr.AddReference('System.Xml')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader, XamlWriter

my_path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
pf_path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86)
sys.path.append(pf_path + '\\IronPython 2.7\\Lib')
import itertools
import math

class MainWindow(Windows.Window):
	def __init__(self, string_xaml):
	
		xr = XmlReader.Create(StringReader(string_xaml))
		self.winLoad = wpf.LoadComponent(self, xr)
		#Get-Set
		self.SelectedLevel = None
		self.SelectedSystemType = None
		self.SelectedPipeType = None
		self.dataRowUnit = None
		#
		self.levels = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
		self.levels.sort(key = lambda lvl : lvl.ProjectElevation)
		#
		self.allSystemType = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).WhereElementIsElementType().ToElements()
		self.allPipeType = FilteredElementCollector(doc).OfClass(PipeType).WhereElementIsElementType().ToElements()
		#
		if sdk_number < 2021:
			lstUnit = [['meter', DisplayUnitType.DUT_METERS, 0.01, 0.05], ['centimeter', DisplayUnitType.DUT_CENTIMETERS, 0.1, 5.0], ['millimeter', DisplayUnitType.DUT_MILLIMETERS, 1, 50.0]]
		else:
			lstUnit = [['meter', Autodesk.Revit.DB.UnitTypeId.Meters, 0.01, 0.05], ['centimeter', Autodesk.Revit.DB.UnitTypeId.Centimeters, 0.1, 5.0], ['millimeter', Autodesk.Revit.DB.UnitTypeId.Millimeters, 1, 50.0]]
		#
		self._tableUnit = DataTable("Unit")
		#self._tableUnit.Rows.Add() # to add a empty row
		#self._tableUnit.AcceptChanges() # to add a empty row
		self._tableUnit.Columns.Add("Name", System.String)
		self._tableUnit.Columns.Add("UnitType", System.Object)
		self._tableUnit.Columns.Add("Accuracy", System.Double)
		self._tableUnit.Columns.Add("MinDistance", System.Double)
		[self._tableUnit.Rows.Add(i, j, k, l) for i, j, k, l in lstUnit]
		# ---NOTE--- send Collection  to ItemSource
		# need to set properties in xaml
		# -> DisplayMemberPath = "Name" 
		# -> SelectedValuePath = "Name" 
		self.Combox_Levels.ItemsSource = self.levels
		self.Combox_SystemType.ItemsSource = self.allSystemType
		self.Combox_PipeType.ItemsSource = self.allPipeType
		self.Combox_Unit.ItemsSource  = self._tableUnit.AsDataView()
		#
		
		
	def Combox_LevelsChanged(self,sender,e):
		#print(sender)
		self.SelectedLevel = sender.SelectedItem
		
	def Combox_SystemTypeChanged(self,sender,e):
		#print(sender)
		self.SelectedSystemType = sender.SelectedItem
		
	def Combox_PipeTypeChanged(self,sender,e):
		#print(sender)
		self.SelectedPipeType = sender.SelectedItem
		
	def Combox_UnitChanged(self,sender,e):
		#print(sender)
		#print(sender.SelectedItem["UnitType"])
		self.dataRowUnit = sender.SelectedItem

	def ButtonOKClick(self,sender,e):
		if self.dataRowUnit is not None : 
			self.Close()
		
	def ButtonCancelClick(self,sender,e):
		self.SelectedLevel = None
		self.SelectedSystemType = None
		self.SelectedPipeType = None
		self.Close()
		
		
def setProjectUnit(dataRowUnit):
	"""set the project same the select """
	TransactionManager.Instance.EnsureInTransaction(doc)
	unit = doc.GetUnits()
	if sdk_number < 2021:
		format = FormatOptions(dataRowUnit["UnitType"])
		format.Accuracy = dataRowUnit["Accuracy"]
		unit.SetFormatOptions(UnitType.UT_Length,format)
	else:
		format = FormatOptions(dataRowUnit["UnitType"])	
		format.Accuracy = dataRowUnit["Accuracy"]
		unit.SetFormatOptions(SpecTypeId.Length,format)
	doc.SetUnits(unit)
	TransactionManager.Instance.TransactionTaskDone()
	


string_xaml = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	x:Name="MainWindow"
	Title="Main"
	MinHeight="420"
	MinWidth="380"
	Width="387"
	Height="431">
	<Grid>
		<Label
			x:Name="Label_Level"
			Content="Select Levels"
			Height="27"
			Width="244"
			Margin="30,18,0,0"
			VerticalAlignment="Top"
			HorizontalAlignment="Left"
			Grid.Row="0"
			Grid.Column="0" />
		<ComboBox
			x:Name="Combox_Levels"
			DisplayMemberPath="Name"
			SelectedValuePath="Name"
			SelectionChanged="Combox_LevelsChanged"
			Grid.Column="0"
			Grid.Row="0"
			VerticalAlignment="Top"
			Height="20"
			Margin="30,53,31,0" />
		<Label
			x:Name="Label_SystemType"
			Grid.Column="0"
			Grid.Row="0"
			HorizontalAlignment="Left"
			VerticalAlignment="Top"
			Margin="30,95,0,0"
			Width="244"
			Height="27"
			Content="Select System Type" />
		<ComboBox
			x:Name="Combox_SystemType"
			DisplayMemberPath="Name"
			SelectedValuePath="Name"
			SelectionChanged="Combox_SystemTypeChanged"
			Grid.Column="0"
			Grid.Row="0"
			VerticalAlignment="Top"
			Height="20"
			Margin="30,130,31,0" />
		<Label
			x:Name="Label_PipeType"
			Content="Select Pipe Type"
			Grid.Column="0"
			Grid.Row="0"
			HorizontalAlignment="Left"
			VerticalAlignment="Top"
			Margin="30,171,0,0"
			Width="159"
			Height="27" />
		<Button
			Grid.Column="0"
			Grid.Row="0"
			x:Name="ButtonOK"
			Content="Continue"
			Click="ButtonOKClick"
			Width="75"
			HorizontalAlignment="Right"
			Height="23"
			VerticalAlignment="Bottom"
			Margin="0,0,31,26" />
		<Button
			Grid.Column="0"
			Grid.Row="0"
			HorizontalAlignment="Left"
			VerticalAlignment="Bottom"
			Margin="30,0,0,26"
			Width="75"
			Height="23"
			x:Name="ButtonCancel"
			Content="Cancel"
			Click="ButtonCancelClick" />
		<ComboBox
			x:Name="Combox_PipeType"
			SelectionChanged="Combox_PipeTypeChanged"
			DisplayMemberPath="Name"
			SelectedValuePath="Name"
			Grid.Column="0"
			Grid.Row="0"
			HorizontalAlignment="Stretch"
			VerticalAlignment="Top"
			Margin="30,206,188,0"
			Height="20" />
		<Label
			x:Name="Label_dwgUnit"
			Grid.Column="0"
			Grid.Row="0"
			HorizontalAlignment="Left"
			VerticalAlignment="Top"
			Margin="30,243,0,0"
			Width="159"
			Height="27"
			Content="Select dwg Unit" />
		<ComboBox
			Grid.Column="0"
			Grid.Row="0"
			VerticalAlignment="Top"
			Height="20"
			Margin="30,278,188,0"
			x:Name="Combox_Unit"
			DisplayMemberPath="Name"
			SelectedValuePath="Name"
			SelectionChanged="Combox_UnitChanged" />
		<Grid.ColumnDefinitions></Grid.ColumnDefinitions>
	</Grid>
</Window>
'''
#
objWpf = MainWindow(string_xaml)
objWpf.ShowDialog()
if objWpf.dataRowUnit is not None : 
	setProjectUnit(objWpf.dataRowUnit)

OUT = objWpf
