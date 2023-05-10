VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitTable 
   Caption         =   "Split Table by Columns"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250.001
   OleObjectBlob   =   "frmSplitTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IView

Private WithEvents mViewModel As SplitTableViewModel
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub chkDeleteExistingSheets_Click()
    mViewModel.DeleteExistingSheets = Me.chkDeleteExistingSheets.Value
End Sub

Private Sub chkShowUnsuitableColumns_Click()
    mViewModel.ShowUnsuitableColumns = Me.chkShowUnsuitableColumns.Value
End Sub

Private Sub chkRemoveOtherSheets_Click()
    mViewModel.RemoveOtherSheets = Me.chkRemoveOtherSheets.Value
End Sub

Private Sub chkShowHiddenColumns_Click()
    mViewModel.ShowHiddenColumns = Me.chkShowHiddenColumns.Value
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub cmbSelectAll_Click()
    'ListViewHelpers.CheckAllItems Me.lvUsedValues
    mViewModel.DoCheckAllTargetSheets
End Sub

Private Sub cmbSelectNone_Click()
    'ListViewHelpers.UncheckAllItems Me.lvUsedValues
    mViewModel.DoCheckNoTargetSheets
End Sub

Private Sub lvAvailableColumns_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.ForeColor <> vbGrayText Then
        mViewModel.SelectListColumnByName Item.Text
    Else
        Item.Checked = False
    End If
End Sub

Private Sub lvAvailableColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.ForeColor <> vbGrayText Then
        mViewModel.SelectListColumnByName Item.Text
    Else
        Item.Selected = False
    End If
End Sub

Private Sub lvUsedValues_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TryCheckTargetSheet Item.Text, Item.Checked
End Sub

Private Sub refColumn_Change()
    'TrySelectColumnFromRefEdit Me.refColumn
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    
    InitalizeFromViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub mViewModel_PropertyChanged(PropertyName As String)
    Select Case PropertyName
        Case "SelectedListObject":
            UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
        Case "UpdateListColumns":
            UpdateAvailableColumns Me.lvAvailableColumns, mViewModel.AvailableColumns
        Case "SelectedListColumn":
            UpdateSelectedColumn Me.lvAvailableColumns, Me.refColumn
        Case "UpdateTargetSheets":
            UpdateTargetSheets Me.lvUsedValues, mViewModel.TargetSheets
        Case "ShowHiddenColumns":
            UpdateAvailableColumns Me.lvAvailableColumns, mViewModel.AvailableColumns
        Case "ShowUnsuitableColumns":
            UpdateAvailableColumns Me.lvAvailableColumns, mViewModel.AvailableColumns
    End Select
    
    UpdateControls
End Sub

Private Sub InitalizeFromViewModel()
    LoadListObjectsToCombobox Me.cboTable, mViewModel.ListObjects
    'UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
    mViewModel_PropertyChanged "SelectedListObject"
    mViewModel_PropertyChanged "UpdateListColumns"
End Sub

Private Sub LoadListObjectsToCombobox(ByVal ComboBox As ComboBox, ByVal ListObjects As Collection)
    ComboBox.Clear
    Dim ListObject As ListObject
    For Each ListObject In ListObjects
        ComboBox.AddItem ListObject.Name
    Next ListObject
End Sub

Private Sub UpdateSelectedListObject(ByVal ComboBox As ComboBox, ByVal ListObject As ListObject)
    ComboBox = ListObject.Name
End Sub

Private Sub UpdateAvailableColumns(ByVal ListView As ListView, ByVal AvailableColumns As AvailableColumns)
    InitalizeAvailableColumns ListView

    Dim ListItem As ListItem
    Dim ThisAvailableColumn As AvailableColumn
    For Each ThisAvailableColumn In AvailableColumns.GetAvailableColumns
        Set ListItem = ListView.ListItems.Add(Text:=ThisAvailableColumn.Name)
        If ThisAvailableColumn.Suitable Then
            ListItem.ListSubItems.Add Text:="Text"
            ListItem.ListSubItems.Add Text:=ThisAvailableColumn.UniqueValueCount
        Else
            ListItem.ListSubItems.Add Text:="Non-text"
            ListItem.ForeColor = vbGrayText
            ListItem.ListSubItems.Item(1).ForeColor = vbGrayText
        End If
    Next ThisAvailableColumn
End Sub

Private Sub InitalizeAvailableColumns(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Column Name", Width:=75
        .ColumnHeaders.Add Text:="Type", Width:=50
        .ColumnHeaders.Add Text:="Values", Width:=50
    
        .View = lvwReport
        .BorderStyle = ccNone
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
    End With
End Sub

Private Sub cboTable_Change()
    mViewModel.SelectListObjectByName Me.cboTable.Text
End Sub

Private Sub UpdateSelectedColumn(ByVal ListView As ListView, ByVal RefEdit As Object)
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        ListItem.Checked = (mViewModel.SelectedListColumn.Name = ListItem.Text)
        ListItem.Selected = (mViewModel.SelectedListColumn.Name = ListItem.Text)
    Next ListItem
    
    'RefEdit.Text = mViewModel.SelectedListColumn.Name
End Sub

Private Sub UpdateTargetSheets(ByVal ListView As ListView, ByVal TargetSheets As Collection)
    InitalizeTargetSheets ListView
    
    Dim ListItem As ListItem
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In TargetSheets
        Set ListItem = ListView.ListItems.Add(Text:=TargetSheet.Name)
        ListItem.Checked = TargetSheet.Used
    Next TargetSheet
End Sub

Private Sub InitalizeTargetSheets(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Sheet Name", Width:=ListView.Width - 16
    
        .View = lvwReport
        .BorderStyle = ccNone
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
    End With
End Sub

Private Sub UpdateControls()
    Me.cmbSelectAll.Enabled = mViewModel.CanSelectAll
    Me.cmbSelectNone.Enabled = mViewModel.CanSelectNone
    Me.cmbOK.Enabled = mViewModel.IsValid
    
    Me.chkDeleteExistingSheets.Value = mViewModel.DeleteExistingSheets
    Me.chkShowUnsuitableColumns.Value = mViewModel.ShowUnsuitableColumns
    Me.chkRemoveOtherSheets.Value = mViewModel.RemoveOtherSheets
    Me.chkShowHiddenColumns.Value = mViewModel.ShowHiddenColumns
End Sub
