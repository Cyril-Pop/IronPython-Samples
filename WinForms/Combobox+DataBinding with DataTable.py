
import clr
import sys
import System

#import Revit API
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *
import Autodesk.Revit.DB as DB

clr.AddReference('System.Data')
from System.Data import *

#import transactionManager and DocumentManager (RevitServices is specific to Dynamo)
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitNodes')
import Revit

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
import System.Drawing
import System.Windows.Forms
from System.Drawing import *
from System.Windows.Forms import *

class MainForm(Form):
	def __init__(self, lst_Elems, title="Title"):
		self._lst_Elems = lst_Elems
		self._title = title
		self._table = DataTable("Elements")
		self._table.Rows.Add() # to add a empty row
		self._table.AcceptChanges() # to add a empty row
		self._table.Columns.Add("Category", System.String)
		self._table.Columns.Add("Name", System.String)
		self._table.Columns.Add("Element", DB.Element)
		# add an concat Column
		self._table.Columns.Add('CustomName', System.String, "Category + ' : ' + Name")
		# populate dataTable
		[self._table.Rows.Add(elem.Category.Name, Element.Name.GetValue(elem), elem) for elem in self._lst_Elems]
		self.choice = None
		
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._comboBox = System.Windows.Forms.ComboBox()
		self._buttonCancel = System.Windows.Forms.Button()
		self._buttonOK = System.Windows.Forms.Button()
		self.SuspendLayout()
		# 
		# comboBox
		# 
		self._comboBox.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
		self._comboBox.Location = System.Drawing.Point(40, 50)
		self._comboBox.Size = System.Drawing.Size(280, 21)
		self._comboBox.DataSource = self._table 
		self._comboBox.DisplayMember = "CustomName"
		self._comboBox.ValueMember = "Element"
		self._comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._comboBox.SelectedIndexChanged += self.ComboBox1SelectedIndexChanged
		# 
		# buttonCancel
		# 
		self._buttonCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._buttonCancel.Location = System.Drawing.Point(240, 122)
		self._buttonCancel.Name = "buttonCancel"
		self._buttonCancel.Size = System.Drawing.Size(80, 25)
		self._buttonCancel.TabIndex = 1
		self._buttonCancel.Text = "Cancel"
		self._buttonCancel.UseVisualStyleBackColor = True
		self._buttonCancel.Click += self.ButtonCancelClick
		# 
		# buttonOK
		# 
		self._buttonOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
		self._buttonOK.Location = System.Drawing.Point(135, 122)
		self._buttonOK.Name = "buttonOK"
		self._buttonOK.Size = System.Drawing.Size(80, 25)
		self._buttonOK.TabIndex = 1
		self._buttonOK.Text = "OK"
		self._buttonOK.UseVisualStyleBackColor = True
		self._buttonOK.Click += self.ButtonOKClick
	
		self.ClientSize = System.Drawing.Size(350, 170)
		self.MinimumSize = self.ClientSize + System.Drawing.Size(20, 20)
		self.Controls.Add(self._comboBox)
		self.Controls.Add(self._buttonOK)
		self.Controls.Add(self._buttonCancel)
		self.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
		self.Name = "MainForm"
		self.Text = self._title
		# set Icon :)
		base64String = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsTAAALEwEAmpwYAAABa0lEQVR4nGNgwAL4JxcYiS+vrxBf19Qvvq6xD8Tmn5JvwEAIiMwpMpTY2LJR7vDEXwonp/9HxiAxiQ3NG0Tnl2A3SHhBqZ/03t6n6BrRsfS+/ifCS8t9MWyWJkIzsiEoLpHY0LwJWYHyqZn/Zz47///t72//3/z69n/Gs/NgMWQ1Ehua18MDTA7Nz62Pjv1//vPL/5Rb28EYxG55eBTFANnDk34KTCrUZwCFMLoTj398+r/14TE4P+/O7v/Tn53D8IroivoyBom1zRPQJS59efW/8v5BgmEhvrZxAgOIINuAdY19DKLL68rJ9sKqulIGUArDCMSHkEBMvbUdjF9gC8SDE34JTM3Tg0XjBvRoBEUdKApBGGQ7RjSub14HTwei80sMpPf1PSE6Ie3teSI6t1wfNSkvLfcFpTBiNIssrvLGmh9ApoJSGCiRoGsE+VliXfN6sdklEH/jA6DAASUScFZe19gHYsMDDA0AAABT+jvUGOyKAAAAAElFTkSuQmCC"
		picIcon = System.Convert.FromBase64String(base64String)
		bmp = System.Drawing.Image.FromStream(System.IO.MemoryStream(picIcon))
		thumb = bmp.GetThumbnailImage(16, 16, bmp.GetThumbnailImageAbort, System.IntPtr.Zero)
		thumb.MakeTransparent()
		self.iconB64 = Icon.FromHandle(thumb.GetHicon())
		self.Icon = self.iconB64
		self.ResumeLayout(False)
		
	def ComboBox1SelectedIndexChanged(self, sender, e):
		if isinstance(sender.SelectedItem['Element'], System.DBNull):
			self.choice = None
		else:
			self.choice = sender.SelectedItem['Element']

	def ButtonOKClick(self, sender, e):
		self.Close()

	def ButtonCancelClick(self, sender, e):
		self.choice = None
		self.Close()

  
lst_rvt_Element = UnwrapElement(IN[0])
objForm = MainForm(lst_rvt_Element)
objForm.ShowDialog()
		
OUT = objForm.choice
