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
'@PredeclaredId
Option Explicit

Private Type TState
    Item As Collection
    Selected As Variant
    ShowHidden As Boolean
    ShowUnsuitable As Boolean
End Type
Private This As TState

Public Property Get Selected() As Variant
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
    CollectionHelpers.CollectionClear This.Item
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        This.Item.Add Item:=AvailableColumn.Create(ListColumn), Key:=ListColumn.Name
    Next ListColumn
End Sub

Public Function GetAvailableColumns() As Collection
    Set GetAvailableColumns = New Collection
    
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        If This.ShowHidden = True Or AvailableColumn.Hidden = False Then
            If This.ShowUnsuitable = True Or AvailableColumn.Suitable = True Then ' TODO CHK
                GetAvailableColumns.Add Item:=AvailableColumn, Key:=AvailableColumn.Name
            End If
        End If
    Next AvailableColumn
End Function

Public Function GetByName(ByVal ColumnName As String) As AvailableColumn
    Set GetByName = This.Item(ColumnName)
End Function

Public Function TrySelectByName(ByVal ColumnName As String) As Boolean
    Dim AvailableColumn As AvailableColumn
    For Each AvailableColumn In This.Item
        AvailableColumn.Selected = AvailableColumn.Name = ColumnName
        If AvailableColumn.Selected Then
            Set This.Selected = AvailableColumn
        End If
    Next AvailableColumn
    
    TrySelectByName = True
End Function
