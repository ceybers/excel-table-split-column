VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model"
Option Explicit

Private Const KEY_PREFIX As String = "K"
Private Const SUITABLE_COLUMN As String = "Text"
Private Const UNSUITABLE_COLUMN As String = "Non-text"

Private Type TState
    ListObject As ListObject
    Item As Collection
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

Public Property Let ShowHidden(ByVal vNewvalue As Boolean)
    This.ShowHidden = vNewvalue
End Property

Public Property Get ShowUnsuitable() As Boolean
    ShowUnsuitable = This.ShowUnsuitable
End Property

Public Property Let ShowUnsuitable(ByVal vNewvalue As Boolean)
    This.ShowUnsuitable = vNewvalue
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

'@Description "Returns a Collection of type AvailableColumn"
Public Function GetAvailableColumns() As Collection
Attribute GetAvailableColumns.VB_Description = "Returns a Collection of type AvailableColumn"
    Dim n As Long
    n = 1
    
    Set GetAvailableColumns = New Collection
    
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        If This.ShowHidden = True Or AvailableColumn.Hidden = False Then
            If This.ShowUnsuitable = True Or AvailableColumn.Suitable = True Then ' TODO CHK
                GetAvailableColumns.Add Item:=AvailableColumn, Key:=AvailableColumn.Name
                AvailableColumn.Index = n
                n = n + 1
            End If
        End If
    Next AvailableColumn
End Function

Public Function GetByName(ByVal ColumnName As String) As AvailableColumn
    Set GetByName = This.Item(ColumnName)
End Function

Public Function TrySelectByName(ByVal ColumnName As String) As Boolean
    If Not GetByName(ColumnName).Suitable Then
        Exit Function
    End If
    
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        AvailableColumn.Selected = (AvailableColumn.Name = ColumnName)
        If AvailableColumn.Selected Then
            Set This.Selected = This.ListObject.ListColumns(AvailableColumn.Name)
            TrySelectByName = True
        End If
    Next AvailableColumn
End Function

Public Sub UpdateListView(ByVal ListView As ListView4)
    Dim ColumnsToDraw As Collection
    Set ColumnsToDraw = GetAvailableColumns
    
    Dim ListItem As ListItem
    For Each AvailableColumn In ColumnsToDraw
        If Not ExistsInCollection(ListView.ListItems, AvailableColumn.Name) Then
            Set ListItem = ListView.ListItems.Add(Text:=AvailableColumn.Name, Key:=KEY_PREFIX & AvailableColumn.Name)
            ListItem.ListSubItems.Add Index:=1, Key:="ColType", Text:=UNSUITABLE_COLUMN
            ListItem.ListSubItems.Add Index:=2, Key:="Count", Text:=""
            ListItem.ListSubItems.Add Index:=3, Key:="SortIndex", Text:="c"
            
            If AvailableColumn.Suitable Then
                ListItem.ListSubItems.Item(1).Text = SUITABLE_COLUMN
                ListItem.ListSubItems.Item(2).Text = AvailableColumn.UniqueValueCount
            Else
                ListItem.ForeColor = vbGrayText
                ListItem.ListSubItems.Item(1).ForeColor = vbGrayText
            End If
        End If
    Next AvailableColumn
    
    Dim i As Long
    For i = ListView.ListItems.Count To 1 Step -1
        If Not ExistsInCollection(ColumnsToDraw, ListView.ListItems.Item(i).Text) Then
            ListView.ListItems.Remove i
        End If
    Next i
    
    For Each AvailableColumn In ColumnsToDraw
        Set ListItem = ListView.ListItems(KEY_PREFIX & AvailableColumn.Name)
        ListItem.Checked = AvailableColumn.Selected
        ListItem.Selected = AvailableColumn.Selected
        ListItem.ListSubItems.Item("SortIndex").Text = AvailableColumn.Index
    Next AvailableColumn
    
    ListView.SortKey = 3
    ListView.SortOrder = lvwAscending
    ListView.Sorted = True
End Sub

Public Sub InitializeListView(ByVal ListView As ListView4)
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
