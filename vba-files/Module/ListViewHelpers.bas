Attribute VB_Name = "ListViewHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder("Helpers")
Option Explicit

'@Obsolete "ZZZ"
Public Sub CheckAllItems(ByVal ListView As ListView)
    Debug.Assert False
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        ListItem.Checked = True
    Next ListItem
End Sub

'@Obsolete "ZZZ"
Public Sub UncheckAllItems(ByVal ListView As ListView)
    Debug.Assert False
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        ListItem.Checked = False
    Next ListItem
End Sub

'@Obsolete "ZZZ"
Public Function SelectionPercent(ByVal ListView As ListView) As Double
    Debug.Assert False
    Dim TotalListItems As Long
    TotalListItems = ListView.ListItems.Count
    If TotalListItems = 0 Then Exit Function
    
    Dim SelectedListItems As Long
    Dim i As Long
    For i = 1 To TotalListItems
        If ListView.ListItems.Item(i).Checked Then
            SelectedListItems = SelectedListItems + 1
        End If
    Next i
    
    SelectionPercent = SelectedListItems / TotalListItems
End Function
