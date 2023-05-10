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
'@Folder("Model")
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
    Set This.ListColumn = ListColumn
    Set This.Item = ColumnAnalysis.Create(ListColumn).GetTargetSheets
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

Public Sub UpdateListView(ByVal ListView As ListView4)
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
    For i = ListView.ListItems.Count To 1 Step -1
        If Not ExistsInCollection(This.Item, ListView.ListItems.Item(i).Text) Then
            ListView.ListItems.Remove i
        End If
    Next i

    For Each TargetSheet In This.Item
        Set ListItem = ListView.ListItems(KEY_PREFIX & TargetSheet.Name)
        ListItem.Checked = TargetSheet.Used
        ListItem.Selected = False
        'ListItem.Selected = TargetSheet.Used
        'ListItem.ListSubItems.Item("SortIndex").Text = TargetSheet.Index
    Next TargetSheet
End Sub

Public Sub InitalizeTargetSheets(ByVal ListView As ListView4)
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
    Set ThisTargetSheet = This.Item(SheetName)
 
    If Not ThisTargetSheet.Suitable Then Exit Function
    
    ThisTargetSheet.Used = Not ThisTargetSheet.Used
    
    TryCheckByName = True
End Function
