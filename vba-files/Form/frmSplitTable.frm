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


Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Me.Hide
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
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub mViewModel_PropertyChanged(PropertyName As String, vNewValue As Object)
    Select Case PropertyName
        Case "SelectedListObject":
            UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
        Case "UpdateListColumns":
            UpdateAvailableColumns Me.lvAvailableColumns, mViewModel.ListColumns
    End Select
End Sub

Private Sub InitalizeFromViewModel()
    LoadListObjectsToCombobox Me.cboTable, mViewModel.ListObjects
    'UpdateSelectedListObject Me.cboTable, mViewModel.SelectedListObject
    mViewModel_PropertyChanged "SelectedListObject", Nothing
    mViewModel_PropertyChanged "UpdateListColumns", Nothing
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

Private Sub UpdateAvailableColumns(ByVal ListView As ListView, ByVal ListColumns As Collection)
    InitalizeAvailableColumns ListView
    
    Dim ListItem As ListItem
    Dim ColumnAnalysis As ColumnAnalysis
    For Each ColumnAnalysis In ListColumns
        Set ListItem = ListView.ListItems.Add(Text:=ColumnAnalysis.ListColumn.Name)
        ListItem.ListSubItems.Add Text:=VarTypeToString(ColumnAnalysis.ColumnVarType)
        
        If ColumnAnalysis.ColumnVarType = (vbArray + vbString) Then
            ListItem.ListSubItems.Add Text:=ColumnAnalysis.UniqueCount
        Else
            ListItem.ForeColor = vbGrayText
            ListItem.ListSubItems.Item(1).ForeColor = vbGrayText
        End If
    Next ColumnAnalysis
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
