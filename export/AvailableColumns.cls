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
'@Folder "TableSplit.Model"
Option Explicit

Private Const KEY_PREFIX As String = "K"
'Private Const SUITABLE_COLUMN As String = "Text"
'Private Const UNSUITABLE_COLUMN As String = "Non-text"
Private Const NO_SUITABLE_COLUMNS As String = "(No suitable columns)"

Private Type TState
    ListObject As ListObject
    Item As Scripting.Dictionary
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
    Set This.Item = New Scripting.Dictionary
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Log.Message "AvailableColumns.Load", "AvailableColumns", Verbose_Level
    Set This.ListObject = ListObject
    Set This.Selected = Nothing
    
    This.Item.RemoveAll
    
    Dim ThisAvailableColumn As AvailableColumn
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        Log.Message "AvailableColumn.Create((" & ListColumn.Name & ")", "AvailableColumns", Verbose_Level
        Set ThisAvailableColumn = AvailableColumn.Create(ListColumn)
        This.Item.Add Item:=ThisAvailableColumn, Key:=ListColumn.Name
        
        If This.Selected Is Nothing And ThisAvailableColumn.Suitable Then
            ThisAvailableColumn.Selected = True
            Set This.Selected = ThisAvailableColumn.ListColumn
        End If
    Next ListColumn
End Sub

'@Description "Returns a Collection of type AvailableColumn"
Public Function GetFilteredColumns() As Collection
Attribute GetFilteredColumns.VB_Description = "Returns a Collection of type AvailableColumn"
    Set GetFilteredColumns = New Collection
    
    Dim Key As Variant
    For Each Key In This.Item.Keys
        Dim AvailableColumn As AvailableColumn
        Set AvailableColumn = This.Item.Item(Key)
        ' Show hidden columns or it wasn't hidden in the first place
        If This.ShowHidden = True Or AvailableColumn.Hidden = False Then
            ' Show unsuitable columns or it wasn't unsuitable in the first place
            If This.ShowUnsuitable = True Or AvailableColumn.Suitable = True Then
                GetFilteredColumns.Add Item:=AvailableColumn, Key:=AvailableColumn.Name
            End If
        End If
    Next Key
    
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
    
    If Not This.Item.Exists(ColumnName) Then Exit Function
    
    If Not This.Item.Item(ColumnName).Suitable Then
        Exit Function
    End If
    
    Dim Element As Variant
    For Each Element In This.Item.Items
        Dim AvailableColumn As AvailableColumn
        Set AvailableColumn = Element
        AvailableColumn.Selected = (AvailableColumn.Name = ColumnName)
        If AvailableColumn.Selected Then
            Set This.Selected = AvailableColumn.ListColumn
            TrySelectByName = True
        End If
    Next Element
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

Public Sub InitializeListView(ByVal ListView As MSComctlLib.ListView)
    Dim Converter As IValueConverter
    Set Converter = New AvailableColumnToListItem
    Converter.Initialize ListView
End Sub

Public Sub Activate()
    If This.Selected Is Nothing Then Exit Sub
    
    This.Selected.Parent.Parent.Activate
    This.Selected.DataBodyRange.Select
End Sub
