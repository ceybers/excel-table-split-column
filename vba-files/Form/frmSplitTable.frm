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

Private Sub refColumn_Change()
    Me.lvAvailableColumns.ListItems(1).Text = Me.refColumn.Text
End Sub

Private Sub UserForm_Activate()
    InitializeAvailableColumns Me.lvAvailableColumns
    InitializeUsedColumns Me.lvUsedColumns
End Sub

Private Sub InitializeAvailableColumns(ByVal lv As ListView)
    With lv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = False
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With
    
    lv.ListItems.Clear
    
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Column Name"
    lv.ColumnHeaders.Add Text:="Column Type", Width:=50
    lv.ColumnHeaders.Add Text:="Uniqueness", Width:=50
    
    Dim li As ListItem
    Set li = lv.ListItems.Add(Text:="Rank")
    li.ListSubItems.Add Text:="Text"
    li.ListSubItems.Add Text:="90%"
End Sub

Private Sub InitializeUsedColumns(ByVal lv As ListView)
    With lv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = True
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With
    
    lv.ListItems.Clear
    
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Sheet Name", Width:=150
    
    Dim li As ListItem
    Set li = lv.ListItems.Add(Text:="Airport")
End Sub
