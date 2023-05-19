VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableColumnToListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableSplit.Converters"
Option Explicit
Implements IValueConverter

Private Const KEY_PREFIX As String = "K"
Private Const SUITABLE_COLUMN As String = "Text"
Private Const UNSUITABLE_COLUMN As String = "Non-text"
'Private Const NO_SUITABLE_COLUMNS As String = "(No suitable columns)"

Private Sub IValueConverter_Convert(ByVal Control As Object, ByVal Object As Object)
    Debug.Assert TypeOf Control Is ListView
    Debug.Assert TypeOf Object Is AvailableColumn
    
    'Dim ListView As ListView
    'Set ListView = Control
    
    'Dim AvailableColumn As AvailableColumn
    'Set AvailableColumn = Object
    
    AddListItem Control, Object
End Sub

Private Sub AddListItem(ByVal ListView As MSComctlLib.ListView, ByVal AvailableColumn As AvailableColumn)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=AvailableColumn.Name, Key:=KEY_PREFIX & AvailableColumn.Name)
    
    ListItem.ListSubItems.Add Index:=1, Key:="ColType", Text:=UNSUITABLE_COLUMN
    ListItem.ListSubItems.Add Index:=2, Key:="Count", Text:=vbNullString
    ListItem.ListSubItems.Add Index:=3, Key:="SortIndex", Text:="c"
    
    If AvailableColumn.Suitable Then
        ListItem.ListSubItems.Item(1).Text = SUITABLE_COLUMN
        ListItem.ListSubItems.Item(2).Text = AvailableColumn.UniqueValueCount
    Else
        ListItem.ForeColor = vbGrayText
        ListItem.ListSubItems.Item(1).ForeColor = vbGrayText
    End If
End Sub

Private Sub IValueConverter_Initialize(ByVal Control As Object)
    Debug.Assert TypeOf Control Is ListView
    'Dim ListView As ListView
    'Set ListView = Control
    
    InitializeListView Control
    InitializeColumnHeaders Control
End Sub

Private Sub InitializeListView(ByVal ListView As MSComctlLib.ListView)
    With ListView
        .ListItems.Clear
        
        .View = lvwReport
        .BorderStyle = ccNone
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
        .LabelEdit = lvwManual
    End With
End Sub

Private Sub InitializeColumnHeaders(ByVal ListView As MSComctlLib.ListView)
    With ListView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Column Name", Width:=75
        .ColumnHeaders.Add Text:="Type", Width:=50
        .ColumnHeaders.Add Text:="Values", Width:=50
        .ColumnHeaders.Add Text:="idx", Width:=0
    End With
End Sub

