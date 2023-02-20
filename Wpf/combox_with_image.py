import clr	
import sys
import System
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
	
from collections import namedtuple
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows")
import System.Drawing
from System.Drawing import *
import System.Windows.Media
import traceback
	
class DropdownInput(Window):

	LAYOUT = '''
			<Window
				xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
				xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
				xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
				xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
				xmlns:local="clr-namespace:WpfApplication1"
				mc:Ignorable="d" 
				Height="300" 
				Width="300" 
				ResizeMode="NoResize"
				Title="A" 
				WindowStartupLocation="CenterScreen" 
				Topmost="True" 
				SizeToContent="Width">
				<Grid Margin="10,0,10,10">
					<Label x:Name="selection_label" Content="Select Item" HorizontalAlignment="Left" Height="30"
						VerticalAlignment="Top"/>
						<ComboBox x:Name="combo_data" HorizontalAlignment="Left" Margin="0,30,0,0" VerticalAlignment="Top" Width="300" DisplayMemberPath="Key" SelectionChanged="Combox_Changed"/>
						<Button x:Name="button_select" Content="Select" HorizontalAlignment="Left" Height="26" Margin="0,63,0,0" VerticalAlignment="Bottom" Width="300" Click="ButtonClick"/>
					<Image x:Name="img" HorizontalAlignment="Center" Height="80" Margin="10,70,10,10" VerticalAlignment="Top" Width="80" Source="" Stretch="Fill" IsEnabled="True"/>
				</Grid>
			</Window>'''
				
	def __init__(self, title, options, description=None):
		self.selected = None
		self.ui = wpf.LoadComponent(self, StringReader(DropdownInput.LAYOUT))
		self.ui.Title = title
		self.error = None
		#
		self.ui.selection_label.Content = description
		#
		self.ui.combo_data.Items.Clear()
		self.ui.combo_data.ItemsSource = options
	
			
	def ButtonClick(self, sender, e):
		self.selected = self.ui.combo_data.SelectedItem
		self.DialogResult = True
		self.Close()
		
	def Combox_Changed(self, sender, e):
		self.selected = self.ui.combo_data.SelectedItem
		try:
			binaryData  = System.Convert.FromBase64String(self.selected.Base64Img)
			bi = System.Windows.Media.Imaging.BitmapImage()
			bi.BeginInit()
			bi.StreamSource = System.IO.MemoryStream(binaryData)
			bi.EndInit()
			#
			self.ui.img.Source = bi
		except Exception as ex:
			self.error = traceback.format_exc()


keys = IN[0] # ["Google", "Autodesk", "Twitter"]
values = IN[1] # ["https://www.google.com/","https://www.autodesk.fr/", "https://twitter.com/"]
lst_base64Img = IN[2] # ["iVBORw0K....", "iVBORw0K....", "iVBORw0K...."]
description = "Choose a Site"

MyImage = namedtuple('MyImage', ['Key', 'Value', 'Base64Img'])
data = []
for key_, value_, base64String in zip(keys, values, lst_base64Img):
	data.append(MyImage(key_, value_, base64String))

form = DropdownInput('InputForm', data, description=description)
form.ShowDialog()
if form.selected is not None:
	OUT = form.selected.Value
