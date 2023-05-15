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
    Item As Collection ' of type ListObject (NOT AvailableTable!)
    Selected As ListObject
    Workbook As Workbook
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
    Set This.Workbook = Workbook
    LoadTablesIntoCollection
    TryLoadFirstTable
End Sub

Private Sub LoadTablesIntoCollection()
    Dim ThisAvailableTable As AvailableTable
    Dim ListObject As ListObject
    For Each ListObject In ListObjectHelpers.GetAllListObjects(This.Workbook)
        Set ThisAvailableTable = AvailableTable.Create(ListObject)
        If ThisAvailableTable.Suitable Then
            This.Item.Add Item:=ListObject, Key:=ListObject.Name
        End If
    Next ListObject
End Sub

Private Sub TryLoadFirstTable()
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

'@Description "Updates the ComboBox to reflect the selected Table."
Public Sub UpdateCombobox(ByVal ComboBox As ComboBox)
Attribute UpdateCombobox.VB_Description = "Updates the ComboBox to reflect the selected Table."
    ' This.Selected should never be nothing, as we select the first ListObject by default,
    ' and we should not be able to complete Loading if there are zero ListObjects.
    If Not This.Selected Is Nothing Then
        ComboBox.Value = This.Selected.Name
    End If
End Sub

'@Description "Clears all items from the given ComboBox, then adds one item for each ListObject, using its Name property."
Public Sub InitializeCombobox(ByVal ComboBox As ComboBox)
Attribute InitializeCombobox.VB_Description = "Clears all items from the given ComboBox, then adds one item for each ListObject, using its Name property."
    ComboBox.Clear
    Dim ListObject As ListObject
    For Each ListObject In This.Item
        ComboBox.AddItem ListObject.Name
    Next ListObject
End Sub
