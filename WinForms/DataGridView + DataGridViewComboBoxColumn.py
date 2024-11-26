import clr
import sys
import System
from System.Collections.Generic import List

#import Revit API
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *
import Autodesk.Revit.DB as DB

clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

clr.AddReference('System.Data')
from System.Data import *

#import transactionManager and DocumentManager (RevitServices is specific to Dynamo)
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

  
class FormSetWkset2(Form):
    def __init__(self, lst_elems, lst_wkset):
        self._lst_wkset = lst_wkset
        self._lst_elems = lst_elems
        self._set_elemTypeId = set(x.GetTypeId() for x in lst_elems)
        self._lst_elemType = [doc.GetElement(xId) for xId in self._set_elemTypeId if xId != ElementId.InvalidElementId]
        #
        self._wksetTable = doc.GetWorksetTable()
        #
        #sort _lst_elemType by Name	
        self._lst_elemType= sorted(self._lst_elemType, key = lambda x : x.FamilyName)
        #
        self._tableDataType = DataTable("ElementType")
        self._tableDataType.Columns.Add("Element", DB.Element)
        self._tableDataType.Columns.Add("FamilyName", System.String)
        self._tableDataType.Columns.Add("Name", System.String)
        self._tableDataType.Columns.Add("Categorie", System.String)
        # populate dataTable
        for x in self._lst_elemType:
            self._tableDataType.Rows.Add(x, x.FamilyName , Element.Name.GetValue(x), x.Category.Name)
        #
        self._tableDataWkset = DataTable("Wkset")
        #self._tableDataWkset.Rows.Add() # to add a empty row
        #self._tableDataWkset.AcceptChanges() # to add a empty row
        self._tableDataWkset.Columns.Add("Workset", DB.Workset)
        self._tableDataWkset.Columns.Add("Name", System.String)
        
        # populate dataTable
        for x in self._lst_wkset:
            self._tableDataWkset.Rows.Add(x, x.Name) 
        #
        self.pairLst = []
        self.InitializeComponent()
        
    
    def InitializeComponent(self):
        dataGridViewCellStyle11 = System.Windows.Forms.DataGridViewCellStyle()
        self._buttonOK = System.Windows.Forms.Button()
        self._dataGridView1 = System.Windows.Forms.DataGridView()
        self._groupBox1 = System.Windows.Forms.GroupBox()
        self._ComboBoxValue = System.Windows.Forms.DataGridViewComboBoxColumn()
        self._dataGridView1.BeginInit()
        self._groupBox1.SuspendLayout()
        self.SuspendLayout()
        #
        self.Shown += self.Form1_Shown
        # 
        # buttonOK
        # 
        self._buttonOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._buttonOK.Location = System.Drawing.Point(1090, 480)
        self._buttonOK.Name = "buttonOK"
        self._buttonOK.Size = System.Drawing.Size(95, 35)
        self._buttonOK.TabIndex = 2
        self._buttonOK.Text = "OK"
        self._buttonOK.UseVisualStyleBackColor = True
        self._buttonOK.Click += self.ButtonOKClick
        # 
        # dataGridView1
        # 
        # set Style
        dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control
        dataGridViewCellStyle11.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
        dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        dataGridViewCellStyle11.WrapMode = getattr(System.Windows.Forms.DataGridViewTriState, "True")
        #
        self._dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle11
        self._dataGridView1.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        self._dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
        self._dataGridView1.AllowUserToAddRows = False
        #
        self._dataGridView1.Location = System.Drawing.Point(6, 28)
        self._dataGridView1.Name = "dataGridView1"
        self._dataGridView1.Size = System.Drawing.Size(1160, 420)
        self._dataGridView1.TabIndex = 3
        self._dataGridView1.DataSource = self._tableDataType
        #
        self._dataGridView1.DataError += self.DataGridViewError
        # 
        # _ComboBoxValue
        # 
        self._ComboBoxValue.HeaderText = "▼ Choice Workset ▼"
        self._ComboBoxValue.Name = "Workset"
        self._ComboBoxValue.Width = 200
        # 
        # groupBox1
        # 
        self._groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._groupBox1.Controls.Add(self._dataGridView1)
        self._groupBox1.Location = System.Drawing.Point(12, 12)
        self._groupBox1.Name = "groupBox1"
        self._groupBox1.Size = System.Drawing.Size(1180, 460)
        self._groupBox1.TabIndex = 4
        self._groupBox1.TabStop = False
        self._groupBox1.Text = "Select Workset"
        # 
        # Form27
        # 
        self.ClientSize = System.Drawing.Size(1200, 530)
        self.MinimumSize = self.ClientSize + System.Drawing.Size(20, 20)
        self.Controls.Add(self._groupBox1)
        self.Controls.Add(self._buttonOK)
        self.Name = "Form27"
        self.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        self.Text = "Title"
        self._dataGridView1.EndInit()
        self._groupBox1.ResumeLayout(False)
        self.ResumeLayout(False)
        
    def Form1_Shown(self, sender, e):
        self.SetCellComboBoxItems()
        
    def SetCellComboBoxItems(self):
        self._dataGridView1.Columns.Add(self._ComboBoxValue)
        for i in range(self._dataGridView1.Rows.Count):
            dgvcbc = self._dataGridView1.Rows[i].Cells[4]
            dgvcbc.DataSource = self._tableDataWkset.Copy()
            dgvcbc.DisplayMember = "Name"
        #
        # hide 1st column
        self._dataGridView1.Columns["Element"].Visible = False
        for idx, col in enumerate(self._dataGridView1.Columns):
            if idx >= 3:
                col.Width = 200
            else:
                col.Width = 250

    def DataGridViewError(self, sender, e):
        print(e.Exception )

    def ButtonOKClick(self, sender, e):
        self.pairLst = []
        for i in range(self._dataGridView1.Rows.Count):
            #
            elem_symbol = self._dataGridView1.Rows[i].Cells[0].Value
            wkset_name = self._dataGridView1.Rows[i].Cells[4].Value
            if elem_symbol is not None and wkset_name is not None:
                # search element in DataTable
                strDataExpression = "[Name] = '" + wkset_name + "'"
                filterP = System.Predicate[System.Object](lambda x : x is not None)
                dtRowA = System.Array.Find(self._tableDataWkset.Select(strDataExpression), filterP)
                wkset = dtRowA["Workset"]
                # remove if None
                if elem_symbol is not None :
                    self.pairLst.append([elem_symbol, wkset])
        self.Close()

            
lst_Elements = UnwrapElement(IN[0])
lst_Wkset = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

form = FormSetWkset2(lst_Elements, lst_Wkset)
form.ShowDialog()

OUT = form.pairLst
