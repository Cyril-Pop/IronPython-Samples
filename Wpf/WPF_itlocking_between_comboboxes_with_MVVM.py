import clr
import sys
import System
from System.Collections.Generic import List, Dictionary
from System.Collections.ObjectModel import ObservableCollection
clr.AddReference("System.Core")
clr.ImportExtensions(System.Linq)

clr.AddReference("System.Xml")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Xml")
clr.AddReference("PresentationCore")
clr.AddReference("System.Windows")
import System.Windows.Controls 
from System.Windows.Controls import *
import System.Windows.Controls.Primitives 
from System.Collections.Generic import List
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows import LogicalTreeHelper 
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

clr.AddReference("IronPython.Wpf")
import wpf

class ViewModel(INotifyPropertyChanged):
    def __init__(self, lstItems):
        super().__init__()
        self._lstItems = ObservableCollection[System.String](lstItems)
        try:
            self.PropertyChanged = None
        except:
            pass
        self._txtValue = ""
        self._itemA = self._lstItems[0] if lstItems else None
        self._itemB = self._lstItems[1] if len(lstItems) > 1 else None
        self._property_changed_handlers = []
        
    # define getter and setter
    
    @property
    def Items(self):
        return self._lstItems
    @Items.setter
    def Items(self, value):
        self._lstItems = value
        self.OnPropertyChanged("Items")
    
    @property
    def TxtValue(self):
        return self._txtValue
    @TxtValue.setter
    def TxtValue(self, value):
        self._txtValue = value
        self.OnPropertyChanged("TxtValue")
        
    @property
    def ItemA(self):
        return self._itemA
    @ItemA.setter
    def ItemA(self, value):
        if value == self._itemB: 
            self.ItemB = None
        self._itemA = value
        self.OnPropertyChanged("ItemA")
        
    @property
    def ItemB(self):
        return self._itemB
    @ItemB.setter
    def ItemB(self, value):
        if value == self._itemA :
            self.ItemA = None
        self._itemB = value
        self.OnPropertyChanged("ItemB")
    
    def OnPropertyChanged(self, property_name):
        event_args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, event_args)

    # Implementation of add/remove_PropertyChanged
    def add_PropertyChanged(self, handler):
        #print(handler)
        if handler not in self._property_changed_handlers:
            self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        #print(handler)
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
                       

class MainForm(Window):
    string_xaml = '''
        <Window 
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                 Title="IronPython WPF Form" Height="200" Width="300">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
        
                    <!-- TextBox -->
                    <Label Grid.Row="0" Grid.Column="0" Content="Enter Text:" VerticalAlignment="Center"/>
                    <TextBox Grid.Row="0" Grid.Column="1" 
                        Name="InputTextBox" Margin="5" VerticalAlignment="Center"
                        Text="{Binding TxtValue, UpdateSourceTrigger=PropertyChanged}" />
        
                    <!-- ComboBox 1 -->
                    <Label Grid.Row="1" Grid.Column="0" Content="Select Option 1:" VerticalAlignment="Center"/>
                    <ComboBox Grid.Row="1" Grid.Column="1" 
                        Name="ComboBox1" Margin="5" VerticalAlignment="Center"
                        ItemsSource="{Binding Items}"
                        SelectedItem="{Binding ItemA, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
        
                    <!-- ComboBox 2 -->
                    <Label Grid.Row="2" Grid.Column="0" Content="Select Option 2:" VerticalAlignment="Center"/>
                    <ComboBox Grid.Row="2" Grid.Column="1" 
                        Name="ComboBox2" Margin="5" VerticalAlignment="Center"
                        ItemsSource="{Binding Items}"
                        SelectedItem="{Binding ItemB, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
        
                    <!-- Submit Button -->
                    <Button Grid.Row="3" Grid.ColumnSpan="2" Content="Submit" 
                            Name="SubmitButton" Margin="5" Width="80" Height="30" 
                            VerticalAlignment="Bottom" HorizontalAlignment="Center" 
                            Click="ButtonClick"/>
                </Grid>
        </Window>'''
  
    def __init__(self, lstvalue):
        super().__init__()
        self.lstvalue = lstvalue
        wpf.LoadComponent(self, StringReader(MainForm.string_xaml))
        # out data
        self.vm = ViewModel(lstvalue)
        self.DataContext = self.vm
        
    def ButtonClick(self, sender, e):
        try:
            self.Close()
        except Exception as ex:
            print(traceback.format_exc())
            
    
lstValues = IN[0]
out = []

my_window = MainForm(lstValues)
my_window.ShowDialog()
out.append(my_window.vm.TxtValue)
out.append(my_window.vm.ItemA)
out.append(my_window.vm.ItemB)

OUT = out
