# Phython-Standard- und DesignScript-Bibliotheken laden
import sys
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
import Autodesk.Revit.DB as DB
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import BindingSource 
clr.AddReference("IronPython.Wpf")
import wpf
import System
from System.Collections.Generic import Dictionary, List
from System import Windows
from System.Windows.Controls import *
from System.IO import StringReader
clr.AddReference('System.Xml')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader, XamlWriter

clr.AddReference("System.Core")
clr.ImportExtensions(System.Linq)

doc = DocumentManager.Instance.CurrentDBDocument

output  = []
class FrmPreviewWindows(Windows.Window):
    def __init__(self):
        InitializeComponent()

class TestWindow(Windows.Window):
    string_xaml = '''
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Window5"
    Width="436"
    Height="537">
    <Grid>
        <Button
            Grid.Column="0"
            Grid.Row="0"
            HorizontalAlignment="Right"
            VerticalAlignment="Bottom"
            Margin="0,0,14,8"
            Width="75"
            Height="23"
            x:Name="Button_ok"
            Content="OK"
            Click="Button_okClick" />
        <ComboBox
            x:Name="comboBoxCategory"
            DisplayMemberPath="Name"
            SelectedValuePath="Name"
            Grid.Row="0"
            VerticalAlignment="Top"
            Height="20"
            Margin="42,75,38,0"
            Grid.Column="0"
            SelectionChanged="comboBoxCategoryChanged" />
        <TextBox
            x:Name="txtBox"
            Grid.Column="0"
            Grid.Row="0"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Bottom"
            Margin="44,0,38,74"
            Height="24" />
        <Label
            x:Name="label_1"
            Grid.Row="0"
            HorizontalAlignment="Left"
            VerticalAlignment="Top"
            Width="122"
            Height="24"
            Content="Select Category"
            Margin="42,43,0,0"
            Grid.Column="0" />
        <Label
            x:Name="label_3"
            Content="Enter Text"
            Grid.Column="0"
            Grid.Row="0"
            HorizontalAlignment="Left"
            VerticalAlignment="Bottom"
            Margin="42,0,0,107"
            Width="122"
            Height="24" />
        <ListView
            x:Name="listViewFamilies"
            Grid.Column="0"
            Grid.Row="0"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Stretch"
            SelectionMode="Extended"
            Margin="42,138,38,152" >
            <ListView.ItemTemplate>
                <DataTemplate>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Label Name="labelFamily" VerticalAlignment="Center"  Margin="0" Content="{Binding listViewFamilies}" Visibility="Hidden" />
                    <CheckBox Name="familyCheck" VerticalAlignment="Center" Margin="0,0,0,0" Content="{Binding Path=Name}" IsChecked="{Binding Path=IsSelected, RelativeSource={RelativeSource FindAncestor, AncestorType={x:Type ListViewItem}}}" />
                </StackPanel>
                </DataTemplate>
            </ListView.ItemTemplate>
        </ListView>
        <Label
            x:Name="label_2"
            Grid.Row="0"
            HorizontalAlignment="Left"
            VerticalAlignment="Top"
            Width="122"
            Height="24"
            Content="Select Families"
            Margin="42,114,0,0"
            Grid.Column="0" />
        <Grid.ColumnDefinitions></Grid.ColumnDefinitions>
    </Grid>
    </Window>'''
  
    def __init__(self):
    
        xr = XmlReader.Create(StringReader(TestWindow.string_xaml))
        self.winLoad = wpf.LoadComponent(self, xr)
        self.catsAllowed = filter(lambda x: x.AllowsBoundParameters,doc.Settings.Categories)
        self.catsAllowed = sorted(self.catsAllowed, key = lambda x : x.Name)
        #Get-Set
        self.SelectedCategroy = None
        self.SelectedFamilies = None
        self.InputText = None
        #
        self.allFams = FilteredElementCollector(doc).OfClass(Family)
        self.allFamTypes = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
        # ---NOTE--- send Collection  to ListView ItemSource
        # need to set properties in xaml
        # -> DisplayMemberPath = "Name" 
        # -> SelectedValuePath = "Name" 
        self.comboBoxCategory.ItemsSource = self.catsAllowed
        
        
    def comboBoxCategoryChanged(self,sender,e):
        self.SelectedCategroy = self.comboBoxCategory.SelectedItem
        self.listViewFamilies.ItemsSource = self.allFams.Where(lambda x: x.FamilyCategory.Id == self.SelectedCategroy.Id)

    def Button_okClick(self,sender,e):
        self.SelectedFamilies = self.listViewFamilies.SelectedItems
        self.InputText = self.txtBox.Text
        self.Close()

try:
    objWpf = TestWindow()
    objWpf.ShowDialog()
    OUT = objWpf.SelectedCategroy, objWpf.SelectedFamilies, objWpf.InputText
except Exception as ex:
    import traceback
    TaskDialog.Show("Error",traceback.format_exc())
    OUT = traceback.format_exc()

