__author__ = "Cyril POUPIN"
__license__ = "MIT license"
__version__ = "1.0.2"

import clr
import sys
import System
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
#import Revit API
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *
import Autodesk.Revit.DB as DB

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

#Get Important vars
doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
sdkNumber = int(app.VersionNumber)

clr.AddReference('System.Data')
from System.Data import *

clr.AddReference("System.Xml")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Xml")
clr.AddReference("PresentationCore")
clr.AddReference("System.Windows")
import System.Windows.Controls 
from System.Windows.Controls import *
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows import LogicalTreeHelper 
from System.Windows.Media import VisualTreeHelper  
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs


import time
import traceback
import itertools


class MainWindow(Window):
    string_xaml = '''
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Selection"
        Height="700" MinHeight="700"
        Width="700" MinWidth="780"
        x:Name="MainWindow">
        <Window.Resources>
        </Window.Resources>
        <Grid Width="auto" Height="auto">
            <Grid.RowDefinitions>
                <RowDefinition Height="30" />
                <RowDefinition />
                <RowDefinition Height="60" />
            </Grid.RowDefinitions>
            <Label
                x:Name="label1"
                Content="Selection"
                Grid.Column="0" Grid.Row="0"
                HorizontalAlignment="Left" VerticalAlignment="Bottom"
                Margin="8,0,366.6,5"
                Width="415" Height="25" />
            <!-- disable on DataGrid Virtualization because we dont use MVVM -->
            <DataGrid
                x:Name="dataGrid"
                AutoGenerateColumns="False"
                CanUserSortColumns="False"
                ItemsSource="{Binding}" 
                Grid.Column="0" Grid.Row="1"
                HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                Margin="8,3,8,7"
                VirtualizingStackPanel.IsVirtualizing="False"
                EnableRowVirtualization="False"
                EnableColumnVirtualization="False"
                CanUserAddRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Family Name" Binding="{Binding FamilyName}" Width="*" />
                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*" />
                    <DataGridTextColumn Header="Category" Binding="{Binding Categorie}" Width="*" />
                    <DataGridTemplateColumn Header="Workset">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <ComboBox x:Name="Combobox"
                                    ItemsSource="{Binding Workset}" 
                                    DisplayMemberPath="Name" 
                                    Width="200"/>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                </DataGrid.Columns>
            </DataGrid>
            <Button
                x:Name="buttonCancel"
                Content="Annuler"
                Grid.Column="0" Grid.Row="2"
                HorizontalAlignment="Left" VerticalAlignment="Bottom"
                Margin="18,13,0,10"
                Height="30" Width="120">
            </Button>
            <Button
                x:Name="buttonOK"
                Content="OK"                
                Grid.Column="0" Grid.Row="2"
                HorizontalAlignment="Right" VerticalAlignment="Bottom"
                Margin="0,12,22,10"
                Height="30" Width="120">
            </Button>
        </Grid>
    </Window>'''
  
    def __init__(self, lst_wkset, lst_elems):
        super().__init__()
        self._lst_wkset = List[DB.Workset](lst_wkset)
        self._lst_elems = lst_elems
        self._set_elemTypeId = set(x.GetTypeId() for x in lst_elems if isinstance(x, FamilyInstance))
        self._lst_elemType = [doc.GetElement(xId) for xId in self._set_elemTypeId if xId != ElementId.InvalidElementId]
        #
        #sort _lst_elemType by Name    
        self._lst_elemType= sorted(self._lst_elemType, key = lambda x : x.FamilyName)
        #
        self._tableDataType = DataTable("ElementType")
        self._tableDataType.Columns.Add("Element", DB.Element)
        self._tableDataType.Columns.Add("FamilyName", System.String)
        self._tableDataType.Columns.Add("Name", System.String)
        self._tableDataType.Columns.Add("Categorie", System.String)
        self._tableDataType.Columns.Add("Workset", List[DB.Workset])
        # populate dataTable
        for x in self._lst_elemType:
            self._tableDataType.Rows.Add(x, x.FamilyName , x.get_Name(), x.Category.Name, self._lst_wkset)
        #
        self.pairLst = []
        #
        xr = XmlReader.Create(StringReader(MainWindow.string_xaml))
        self.winLoad = XamlReader.Load(xr) 
        self.InitializeComponent()
        
    def InitializeComponent(self):
        try:
            self.Content = self.winLoad.Content
            #
            self.dataGrid = LogicalTreeHelper.FindLogicalNode(self.winLoad, "dataGrid")
            #
            self.buttonCancel = LogicalTreeHelper.FindLogicalNode(self.winLoad, "buttonCancel")
            self.buttonCancel.Click += self.ButtonCancelClick
            #
            self.buttonOK = LogicalTreeHelper.FindLogicalNode(self.winLoad, "buttonOK")
            self.buttonOK.Click += self.ButtonOKClick
            #
            self.winLoad.Loaded += self.OnLoad
            #
            self.dataGrid.DataContext = self._tableDataType
        except Exception as ex:
            print(traceback.format_exc())
        
            
    def OnLoad(self, sender, e):
        print("UI loaded")
        
    def ButtonCancelClick(self, sender, e):
        self.outSelection = []
        self.winLoad.Close()
        
    def FindVisualChild(self, parent, child_type):
        """Recursively finds a child of a specific type in the visual tree."""
        num_children = VisualTreeHelper.GetChildrenCount(parent)
        for i in range(num_children):
            child = VisualTreeHelper.GetChild(parent, i)
            if isinstance(child, child_type):
                return child
            result = FindVisualChild(child, child_type)
            if result is not None:
                return result
        return None
        
    def ButtonOKClick(self, sender, e):
        try:
            self.pairLst = []
            for row in self.dataGrid.Items:  # loop through all rows in the DataGrid
                container = self.dataGrid.ItemContainerGenerator.ContainerFromItem(row)  # Get row container
                if container is not None:
                    # locate the ComboBox inside the row
                    cell_content = self.dataGrid.Columns.get_Item(3).GetCellContent(row)
                    combo_box = self.FindVisualChild(cell_content, System.Windows.Controls.ComboBox)
                    print(f"{combo_box=}")
                    if combo_box:
                        selected_item = combo_box.SelectedItem  # extract selected item
                        print(selected_item)
                        self.pairLst.append([row["Element"], selected_item])
            self.winLoad.Close()
        except Exception as ex:
            print(traceback.format_exc())
            
        
lst_Elements = UnwrapElement(IN[0])
lst_Wkset = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
objWindow = MainWindow(lst_Wkset, lst_Elements)
objWindow.winLoad.ShowDialog()

OUT = objWindow.pairLst
