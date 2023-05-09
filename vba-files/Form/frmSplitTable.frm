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

Private ViewModel As SplitTableViewModel
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Hide
End Sub

Private Sub UserForm_Activate()
    'InitializeAvailableColumns Me.lvAvailableColumns
    'InitializeUsedValues Me.lvUsedValues
End Sub

Private Sub InitialiseListObject(ByVal cbo As ComboBox)
    cbo.Clear
    cbo.AddItem "Hello world"
    cbo.Value = "Hello world"
End Sub

Private Sub InitialiseListColumn() 'ByVal re As RefEdit)
    Me.refColumn.Text = "Testing 123"
End Sub

Private Sub InitializeAvailableColumns(ByVal lv As ListView)
    With lv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = False
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .view = lvwReport
    End With
    
    lv.ListItems.Clear
    
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Column Name"
    lv.ColumnHeaders.Add Text:="Column Type", Width:=50
    lv.ColumnHeaders.Add Text:="Uniqueness", Width:=50
    
    Dim ca As ColumnAnalysis
    For Each ca In This.ViewModel.AvailableColumns
        Dim li As ListItem
        Set li = lv.ListItems.Add(Text:=ca.ListColumn.Name)
        li.ListSubItems.Add Text:=VarTypeToString(ca.ColumnVarType - vbArray)
        li.ListSubItems.Add Text:=FormatPercent(ca.Uniqueness, 0)
    Next ca
End Sub

Private Sub InitializeUsedValues(ByVal lv As ListView)
    With lv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = True
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .view = lvwReport
    End With
    
    lv.ListItems.Clear
    
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Sheet Name", Width:=150
    
    Dim li As ListItem
    Set li = lv.ListItems.Add(Text:="Airport")
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

Public Function ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    
    InitialiseListObject Me.cboTable
    InitialiseListColumn
    InitializeAvailableColumns Me.lvAvailableColumns
    InitializeUsedValues Me.lvUsedValues
    
    Show
    ShowDialog = Not This.IsCancelled
End Function
