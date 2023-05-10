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

Private Const SELECT_ALL As String = "(Select all)"

Private Type TState
    Item As Collection ' of type TargetSheet
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Sub Load(ByVal ListColumn As ListColumn)
    Set This.Item = ColumnAnalysis.Create(ListColumn).GetTargetSheets
End Sub

Public Function GetSelectedSheetNames() As Collection
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

Public Sub LoadListView(ByVal ListView As ListView)
    If ListView.ListItems.Count <> This.Item.Count Then
        ListView.ListItems.Clear
    End If
    
    Dim ThisTargetSheet As TargetSheet
    If ListView.ListItems.Count = 0 Then
        For Each ThisTargetSheet In This.Item
            ListView.ListItems.Add Text:=ThisTargetSheet.Name, Key:=ThisTargetSheet.Name
        Next ThisTargetSheet
    End If
    
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        ListItem.Checked = This.Item(ListItem.Text).Used
        ListItem.Selected = This.Item(ListItem.Text).Used
    Next ListItem
End Sub

Public Function TryCheckByName(ByVal SheetName As String, ByVal vNewvalue As Boolean) As Boolean
    If SheetName = SELECT_ALL Then
        'If vNewValue And CanSelectAll Then
        '    SelectAll
        '    TryCheckByName = True
        'ElseIf CanSelectNone Then
        '    SelectNone
        '    TryCheckByName = True
        'End If
    Else
        Dim ThisTargetSheet As TargetSheet
        Set ThisTargetSheet = This.Item(SheetName)
        
        If vNewvalue And ThisTargetSheet.Used = False Then
            ThisTargetSheet.Used = True
            TryCheckByName = True
        ElseIf ThisTargetSheet.Used = True Then
            ThisTargetSheet.Used = False
            TryCheckByName = True
        End If
    End If
End Function
