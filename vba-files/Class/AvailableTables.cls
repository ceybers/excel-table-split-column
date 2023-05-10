VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Model")
Option Explicit

Private Type TState
    Item As Collection
    Selected As ListObject
End Type
Private This As TState

Public Property Get Selected() As ListObject
    Set Selected = This.Selected
End Property

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Sub Load(ByVal Workbook As Workbook)
    CollectionHelpers.CollectionClear This.Item

    Dim ThisAvailableTable As AvailableTable
    Dim ListObject As ListObject
    For Each ListObject In ListObjectHelpers.GetAllListObjects(Workbook)
        Set ThisAvailableTable = AvailableTable.Create(ListObject)
        If ThisAvailableTable.Suitable Then
            This.Item.Add Item:=ListObject, Key:=ThisAvailableTable.Name
        End If
    Next ListObject

    If This.Item.Count > 0 Then
        Set This.Selected = This.Item.Item(1)
    End If
End Sub

Public Function TrySelectByName(ByVal TableName As String) As Boolean
    Dim ListObject As ListObject
    For Each ListObject In This.Item
        If ListObject.Name = TableName Then
            Set This.Selected = ListObject
            TrySelectByName = True
        End If
    Next ListObject
End Function

Public Sub UpdateCombobox(ByVal ComboBox As ComboBox)
    ComboBox.Value = This.Selected.Name
End Sub

Public Sub InitializeCombobox(ByVal ComboBox As ComboBox)
    ComboBox.Clear
    Dim ListObject As ListObject
    For Each ListObject In This.Item
        ComboBox.AddItem ListObject.Name
    Next ListObject
End Sub
