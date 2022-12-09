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
	clr.AddReferenceByPartialName("WindowsBase")
except IOError:
	raise
	
from System.IO import StringReader
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application

try:
	import wpf
	import time
except ImportError:
	raise

class CreateProgressWindow(Window):
	LAYOUT = '''
	<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
		Title="WindowPgb5"
		Height="300"
		Width="300">
		<Grid>
			<ProgressBar
				Grid.Column="0"
				Grid.Row="0"
				HorizontalAlignment="Left"
				VerticalAlignment="Top"
				Margin="40,122,0,0"
				Width="210"
				Height="20"
				x:Name="pbar" />
		</Grid>
	</Window>
	'''
	
	def __init__(self):
		self.ui = wpf.LoadComponent(self, StringReader(CreateProgressWindow.LAYOUT))

	def _dispatch_updater(self):
		# ask WPF dispatcher for gui update
		self.pbar.Dispatcher.Invoke(System.Action(self._update_pbar),
									System.Windows.Threading.DispatcherPriority.Background)


	def _update_pbar(self):
		self.pbar.Value = self.new_value
		if self.pbar.Value == self.pbar.Maximum:
			self.Close()
	
	def update_progress(self, value):
		self.new_value = value
		self._dispatch_updater()

			
pb = CreateProgressWindow()
pb.Show()
for i in range(1, 101):
	time.sleep(0.1)
	pb.update_progress(i)

