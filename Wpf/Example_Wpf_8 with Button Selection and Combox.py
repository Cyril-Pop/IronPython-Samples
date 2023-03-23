import clr	
import sys
import System
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

#import Revit APIUI namespace
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
uiapp=DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
sdk_number = int(app.VersionNumber)

sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib')
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\DLLs')

try:
	clr.AddReference("IronPython.Wpf")
	clr.AddReference('System.Core')
	clr.AddReference('System.Xml')
	clr.AddReference('PresentationCore')
	clr.AddReference('PresentationFramework')
except IOError:
	raise
	
from System.IO import StringReader
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application

from System import Uri
from System.Windows.Media.Imaging import BitmapImage

try:
	import wpf
except ImportError:
	raise
	
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows")
import System.Drawing
from System.Drawing import *
import System.Windows.Media
import traceback

class CustomISelectionFilter(ISelectionFilter):
	def __init__(self, bic_category):
		self.bic_category = bic_category
	def AllowElement(self, e):
		if e.Category.Id == ElementId(self.bic_category):
			return True
		else:
			return False
	def AllowReference(self, ref, point):
		return true
	
class MyWindow(Window):

	LAYOUT = '''
		<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
				xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
				Title="MainWindow" Height="400" Width="600"  WindowStartupLocation="CenterScreen"  ResizeMode="CanResize" MinHeight="400" MinWidth="600">
				<Grid>
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="20" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50"/>
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50"/>
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="*" />
						<ColumnDefinition Width="50" />
						<ColumnDefinition Width="20" />
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height ="20" />
						<RowDefinition Height ="50" />
						<RowDefinition Height ="50" />
						<RowDefinition Height ="50" />
						<RowDefinition Height ="50" />
						<RowDefinition Height ="*" />
						<RowDefinition Height ="50" />
						<RowDefinition Height ="20" />
					</Grid.RowDefinitions>
					<Button  x:Name="btn_Select_Pipe" Content="Select Pipe" Grid.Column="9" Grid.Row="3"  Grid.ColumnSpan="2" Background="LightBlue"  Margin="10,0,10,20" Click="btn_Click_SelectPipe"/>              
					<Label Content="Pipe Arrow Placer " Foreground="#BF301A" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="2" Grid.Row="1"  Grid.ColumnSpan="7" Margin="10,0,10,20"/>
					<Label Content="Select a Pipe from pipe Network " Foreground="#BF301A" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="1" Grid.Row="3"  Grid.ColumnSpan="7" Margin="10,0,10,20"/>
					<Label Content="Select Flow Arrow Family " Foreground="#BF301A" FontWeight="Bold" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="1" Grid.Row="4"  Grid.ColumnSpan="7" Margin="10,0,10,20"/>
					<ComboBox x:Name="ComboBox_AnnoFamily" Grid.Column="9" Grid.Row="4"  Grid.ColumnSpan="2" Background="LightBlue"  Margin="10,0,10,20" DisplayMemberPath="Name" SelectionChanged="Combox_Changed">
					</ComboBox>
		
			</Grid>
				
		</Window>'''
				
	def __init__(self):
		self.AnnoFamily_selected = None
		self.Pipe_Sel = None
		self.ui = wpf.LoadComponent(self, StringReader(MyWindow.LAYOUT))
		Generic_Anno_Famlies = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements()

		#
		self.ui.ComboBox_AnnoFamily.Items.Clear()
		self.ui.ComboBox_AnnoFamily.ItemsSource = Generic_Anno_Famlies
	
			
	def btn_Click_SelectPipe(self, sender, e):
		try:
			self.Hide()
			ref = uidoc.Selection.PickObject(ObjectType.Element,CustomISelectionFilter(BuiltInCategory.OST_PipeCurves),"Select a Pipe")
			self.Pipe_Sel = doc.GetElement(ref)
			self.ShowDialog()
		except Exception as ex:
			print(ex)
			TaskDialog.Show("Operation canceled","Canceled by the user")
			self.Pipe_Sel = None
		
	def Combox_Changed(self, sender, e):
		self.AnnoFamily_selected = self.ui.ComboBox_AnnoFamily.SelectedItem
		print(self.AnnoFamily_selected)
		print(Element.Name.GetValue(self.AnnoFamily_selected))


objWpf = MyWindow()
objWpf.ShowDialog()

OUT = objWpf.AnnoFamily_selected, objWpf.Pipe_Sel
