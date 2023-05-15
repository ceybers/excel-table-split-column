VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Model"
Option Explicit

Private Const KEY_PREFIX As String = "K"
Private Const SUITABLE_COLUMN As String = "Text"
Private Const UNSUITABLE_COLUMN As String = "Non-text"
Private Const NO_SUITABLE_COLUMNS As String = "(No suitable columns)"

Private Type TState
    ListObject As ListObject
    Item As Collection ' of type AvailableColumn
    Selected As ListColumn
    
    ShowHidden As Boolean
    ShowUnsuitable As Boolean
End Type
Private This As TState

Public Property Get Selected() As ListColumn
    Set Selected = This.Selected
End Property

Public Property Get ShowHidden() As Boolean
    ShowHidden = This.ShowHidden
End Property

Public Property Let ShowHidden(ByVal vNewValue As Boolean)
    This.ShowHidden = vNewValue
End Property

Public Property Get ShowUnsuitable() As Boolean
    ShowUnsuitable = This.ShowUnsuitable
End Property

Public Property Let ShowUnsuitable(ByVal vNewValue As Boolean)
    This.ShowUnsuitable = vNewValue
End Property

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject

    CollectionHelpers.CollectionClear This.Item
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        This.Item.Add Item:=AvailableColumn.Create(ListColumn), Key:=ListColumn.Name
    Next ListColumn

    Set This.Selected = ListObject.ListColumns.Item(1)
End Sub

'@Description "Returns a filtered Collection of type AvailableColumn"
Public Function GetFilteredColumns() As Collection
Attribute GetFilteredColumns.VB_Description = "Returns a Collection of type AvailableColumn"
    Set GetFilteredColumns = New Collection
    
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        ' Show hidden columns or it wasn't hidden in the first place
        If This.ShowHidden = True Or AvailableColumn.Hidden = False Then
            ' Show unsuitable columns or it wasn't unsuitable in the first place
            If This.ShowUnsuitable = True Or AvailableColumn.Suitable = True Then
                GetFilteredColumns.Add Item:=AvailableColumn, Key:=AvailableColumn.Name
            End If
        End If
    Next AvailableColumn
    
    ' This index should be matching the index in the ListView when we update it.
    Dim i As Long
    For i = 1 To GetFilteredColumns.Count
        Set AvailableColumn = GetFilteredColumns.Item(i)
        AvailableColumn.Index = i
    Next i
End Function

Public Function TrySelectByName(ByVal ColumnName As String) As Boolean
    If ColumnName = NO_SUITABLE_COLUMNS Then
        Exit Function
    End If
    
    If Not This.Item.Item(ColumnName).Suitable Then
        Exit Function
    End If
    
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        AvailableColumn.Selected = (AvailableColumn.Name = ColumnName)
        If AvailableColumn.Selected Then
            ' Gets the reference to the ListObject because we didn't store it in AvailableColumn
            ' TODO Just store the reference in AvailableColumn
            'Set This.Selected = This.ListObject.ListColumns.Item(AvailableColumn.Name)
            Set This.Selected = AvailableColumn.ListColumn
            TrySelectByName = True
        End If
    Next AvailableColumn
End Function

Public Sub UpdateListView(ByVal ListView As MSComctlLib.ListView)
    Dim ColumnsToDraw As Collection
    Set ColumnsToDraw = GetFilteredColumns
    
    ' Remove obsolete items from ListView if they are not in ColumnsToDraw
    Dim i As Long
    For i = ListView.ListItems.Count To 1 Step -1
        If Not ExistsInCollection(ColumnsToDraw, ListView.ListItems.Item(i).Text) Then
            ListView.ListItems.Remove i
        End If
    Next i
    
    ' This comes last because we need the earlier code to remove obsolete items first.
    If ColumnsToDraw.Count = 0 Then
        UpdateListViewEmpty ListView
        Exit Sub
    End If
    
    ' Add missing items from ColumnsToDraw into ListView
    Dim Converter As IValueConverter
    Set Converter = New AvailableColumnToListItem
    
    Dim ListItem As ListItem
    For Each AvailableColumn In ColumnsToDraw
        If Not ExistsInCollection(ListView.ListItems, AvailableColumn.Name) Then
            Converter.Convert ListView, AvailableColumn
        End If
    Next AvailableColumn
    
    ' Rebuild the SortIndex in ListView
    For Each AvailableColumn In ColumnsToDraw
        Set ListItem = ListView.ListItems.Item(KEY_PREFIX & AvailableColumn.Name)
        ListItem.Checked = AvailableColumn.Selected
        ListItem.Selected = AvailableColumn.Selected
        ListItem.ListSubItems.Item("SortIndex").Text = AvailableColumn.Index
    Next AvailableColumn
    
    ' Sort accordingly
    ListView.SortKey = 3
    ListView.SortOrder = lvwAscending
    ListView.Sorted = True
End Sub

Private Sub UpdateListViewEmpty(ByVal ListView As MSComctlLib.ListView)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=NO_SUITABLE_COLUMNS)
    ListItem.ForeColor = vbGrayText
    
    With ListView
        .CheckBoxes = False
        .ColumnHeaders.Item(1).Width = ListView.Width - 8
        .ColumnHeaders.Item(2).Width = 0
        .ColumnHeaders.Item(3).Width = 0
    End With
End Sub

Public Sub InitializeListView(ByVal ListView As ListView4)
    Dim Converter As IValueConverter
    Set Converter = New AvailableColumnToListItem
    Converter.Initialize ListView
    Exit Sub
    
    With ListView
        .ListItems.Clear
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Column Name", Width:=75
        .ColumnHeaders.Add Text:="Type", Width:=50
        .ColumnHeaders.Add Text:="Values", Width:=50
        .ColumnHeaders.Add Text:="idx", Width:=0
        
        .View = lvwReport
        .BorderStyle = ccNone
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
        .LabelEdit = lvwManual
    End With
End Sub
