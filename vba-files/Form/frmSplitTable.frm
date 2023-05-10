VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitTable 
   Caption         =   "Split Table by Columns"
   ClientHeight    =   6015
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
'@Folder "MVVM"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mViewModel As SplitTableViewModel
Attribute mViewModel.VB_VarHelpID = -1
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
    If mViewModel.TargetSheets.SelectAll Then
        mViewModel_PropertyChanged "UpdateTargetSheets"
    End If
End Sub

Private Sub cmbSelectNone_Click()
    If mViewModel.TargetSheets.SelectNone Then
        mViewModel_PropertyChanged "UpdateTargetSheets"
    End If
End Sub

Private Sub lvAvailableColumns_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Not mViewModel.TrySelectColumnByName(Item.Text) Then
        Item.Checked = False
    End If
End Sub

Private Sub lvAvailableColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TrySelectColumnByName Item.Text
End Sub

Private Sub lvUsedValues_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TryCheckTargetSheet Item.Text, Item.Checked
End Sub

Private Sub lvUsedValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TryCheckTargetSheet Item.Text, True
End Sub

'Private Sub refColumn_Change()
    'TrySelectColumnFromRefEdit Me.refColumn
'End Sub

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
        Case "SelectedListColumn":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "ShowUnsuitableColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "ShowHiddenColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
            
        Case "SelectedListObject":
            UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
        Case "UpdateTargetSheets":
            UpdateTargetSheets Me.lvUsedValues, mViewModel.TargetSheets
    End Select
    
    UpdateControls
End Sub

Private Sub InitalizeFromViewModel()
    
    LoadListObjectsToCombobox Me.cboTable, mViewModel.ListObjects
    'UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
    mViewModel_PropertyChanged "SelectedListObject"
    
    mViewModel.AvailableColumns.InitializeListView Me.lvAvailableColumns
    mViewModel_PropertyChanged "SelectedListColumn"
    mViewModel_PropertyChanged "UpdateTargetSheets"
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

Private Sub cboTable_Change()
    mViewModel.SelectListObjectByName Me.cboTable.Text
End Sub

Private Sub UpdateSelectedColumn(ByVal ListView As ListView, ByVal RefEdit As Object)
    mViewModel.AvailableColumns.UpdateListView ListView
End Sub

' TODO Move this into TargetSheets class module
Private Sub UpdateTargetSheets(ByVal ListView As ListView, ByVal TargetSheets As TargetSheets)
    InitalizeTargetSheets ListView
    TargetSheets.LoadListView ListView
End Sub

' TODO Move this into TargetSheets class module
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
