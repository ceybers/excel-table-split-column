VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "TableSplit.Model"
Option Explicit

Private Const KEY_PREFIX As String = "K"

Private Type TState
    ListColumn As ListColumn
    Item As Collection ' of type TargetSheet
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Sub Load(ByVal ListColumn As ListColumn)
    If Not ListColumn Is Nothing Then
        Set This.ListColumn = ListColumn
        Set This.Item = ColumnAnalysis.Create(ListColumn).GetTargetSheets
    Else
        CollectionHelpers.CollectionClear This.Item
    End If
End Sub

'@Description "Returns a Collection of TargetSheet where the Used property is True"
Public Function GetSelectedSheetNames() As Collection
Attribute GetSelectedSheetNames.VB_Description = "Returns a Collection of TargetSheet where the Used property is True"
    Set GetSelectedSheetNames = New Collection
    
    Dim ThisTargetSheet As TargetSheet
    For Each ThisTargetSheet In This.Item
        If ThisTargetSheet.Used = True Then
            GetSelectedSheetNames.Add ThisTargetSheet.Name
        End If
    Next ThisTargetSheet
End Function

Public Function SelectAll() As Boolean
    If CanSelectAll = False Then Exit Function
    
    Dim ThisTargetSheet As TargetSheet
    For Each ThisTargetSheet In This.Item
        ThisTargetSheet.Used = True
    Next ThisTargetSheet
    
    SelectAll = True
End Function

Public Function SelectNone() As Boolean
    If CanSelectNone = False Then Exit Function

    Dim ThisTargetSheet As TargetSheet
    For Each ThisTargetSheet In This.Item
        ThisTargetSheet.Used = False
    Next ThisTargetSheet
    
    SelectNone = True
End Function

Private Function SelectedItems() As Long
    Dim ThisTargetSheet As TargetSheet
    For Each ThisTargetSheet In This.Item
        If ThisTargetSheet.Used = True Then
            SelectedItems = SelectedItems + 1
        End If
    Next ThisTargetSheet
End Function

Public Property Get CanSelectAll() As Boolean
    CanSelectAll = SelectedItems < This.Item.Count
End Property

Public Property Get CanSelectNone() As Boolean
    CanSelectNone = SelectedItems > 0 And This.Item.Count > 0
End Property

Public Sub UpdateListView(ByVal ListView As MSComctlLib.ListView)
    Dim ListItem As ListItem

    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.Item
        If Not ExistsInCollection(ListView.ListItems, TargetSheet.Name) Then
            Set ListItem = ListView.ListItems.Add(Text:=TargetSheet.Name, Key:=KEY_PREFIX & TargetSheet.Name)
            ListItem.ListSubItems.Add Index:=1, Key:="SortIndex", Text:=0 'TODO NYI
            If Not TargetSheet.Suitable Then
                ListItem.ForeColor = vbGrayText
            End If
        End If
    Next TargetSheet
    
    Dim i As Long
    Dim j As Long
    Dim found As Boolean
    For i = ListView.ListItems.Count To 1 Step -1
        found = False
        Dim LHS As String
        LHS = ListView.ListItems.Item(i).Text
        
        For j = 1 To This.Item.Count
            Set TargetSheet = This.Item.Item(j)
            Dim RHS As String
            RHS = TargetSheet.Name
            If LHS = RHS Then
                found = True
                Exit For
            End If
        Next j
        
        If found = False Then
            ListView.ListItems.Remove i
        End If
    Next i
    
    For Each TargetSheet In This.Item
        If TryGetListItem(ListView, (KEY_PREFIX & TargetSheet.Name), ListItem) Then
            'Set ListItem = ListView.ListItems.Item(KEY_PREFIX & TargetSheet.Name)
            ListItem.Checked = TargetSheet.Used
            ListItem.Selected = False
            'ListItem.Selected = TargetSheet.Used
            'ListItem.ListSubItems.Item("SortIndex").Text = TargetSheet.Index
        End If
    Next TargetSheet
End Sub

Private Function TryGetListItem(ByVal ListView As MSComctlLib.ListView, ByVal Key As String, ByRef OutValue As ListItem) As Boolean
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        If ListItem.Key = Key Then
            Set OutValue = ListItem
            TryGetListItem = True
        End If
    Next ListItem
End Function

Public Sub InitalizeTargetSheets(ByVal ListView As MSComctlLib.ListView)
    With ListView
        .ListItems.Clear
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Sheet Name", Width:=ListView.Width - 16
        .ColumnHeaders.Add Text:="idx", Width:=0
    
        .View = lvwReport
        .BorderStyle = ccNone
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True
        .LabelEdit = lvwManual
        .MultiSelect = False
        '.HotTracking = True
        '.HideSelection = False
    End With
End Sub

Public Function TryCheckByName(ByVal SheetName As String, ByVal vNewValue As Boolean) As Boolean
    Dim ThisTargetSheet As TargetSheet
    Set ThisTargetSheet = This.Item.Item(SheetName)
 
    If Not ThisTargetSheet.Suitable Then Exit Function
    
    ThisTargetSheet.Used = Not ThisTargetSheet.Used
    
    TryCheckByName = True
End Function
