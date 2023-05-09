VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SplitTableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SplitTable"
Option Explicit

Private Type TState
    ListObject As ListObject
    ListColumn As ListColumn
    AvailableColumns As Collection
    UsedValues As Collection
    HideUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
End Type
Private This As TState

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
  End Property
  
'Public Property Set ListObject(ByVal vNewValue As ListObject)
'    Set This.ListObject = vNewValue
'End Property

Public Property Get ListColumn() As ListColumn
    Set ListColumn = This.ListColumn
End Property

'Public Property Set ListColumn(ByVal vNewValue As ListColumn)
'    Set This.Column = vNewValue
'End Property

Public Property Get AvailableColumns() As Collection
    Set AvailableColumns = This.AvailableColumns
End Property

'Public Property Set AvailableColumns(ByVal vNewValue As Collection)
'    Set This.AvailableColumns = vNewValue
'End Property

Public Property Get UsedValues() As Collection
    Set UsedValues = This.UsedValues
    End Property
    
'Public Property Set UsedColumns(ByVal vNewValue As Collection)
'    Set This.UsedColumns = vNewValue
'End Property

Public Property Get HideUnsuitableColumns() As Boolean
    HideUnsuitableColumns = This.HideUnsuitableColumns
End Property

Public Property Let HideUnsuitableColumns(ByVal vNewValue As Boolean)
    This.HideUnsuitableColumns = vNewValue
End Property

Public Property Get DeleteExistingSheets() As Boolean
    DeleteExistingSheets = This.DeleteExistingSheets
End Property

Public Property Let DeleteExistingSheets(ByVal vNewValue As Boolean)
    This.DeleteExistingSheets = vNewValue
End Property
    
Public Property Get RemoveOtherSheets() As Boolean
    RemoveOtherSheets = This.RemoveOtherSheets
End Property

Public Property Let RemoveOtherSheets(ByVal vNewValue As Boolean)
    This.RemoveOtherSheets = vNewValue
End Property

Public Sub Load(ByVal lc As ListColumn)
    Debug.Assert Not lc Is Nothing
    Set This.ListColumn = lc
    Set This.ListObject = lc.Parent
    
    LoadAvailableColumns
    
    LoadUsedValues
End Sub

Private Sub LoadAvailableColumns()
    Debug.Assert This.AvailableColumns Is Nothing
    Set This.AvailableColumns = New Collection
    
    Dim lc As ListColumn
    Dim ca As ColumnAnalysis
    For Each lc In This.ListObject.ListColumns
        Set ca = New ColumnAnalysis
        ca.Analyse lc
        This.AvailableColumns.Add Item:=ca, Key:=lc.Name
    Next
End Sub

Private Sub LoadUsedValues()
    Set This.UsedValues = New Collection
    
    'Dim ca As ColumnAnalysis
    'Set ca = This.AvailableColumns.Item(4)
    
    Dim uv As UsedValue
    Set uv = New UsedValue
    uv.Name = "Test"
    uv.Used = True
    
    This.UsedValues.Add Item:=uv
End Sub
