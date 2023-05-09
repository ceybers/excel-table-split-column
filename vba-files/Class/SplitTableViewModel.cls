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
    Table As ListObject
    Column As ListColumn
    AvailableColumns As Collection
    UsedColumns As Collection
    HideUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
End Type
Private This As TState

Public Property Get Table() As ListObject
    Set Table = This.Table
  End Property
  
Public Property Let Table(ByVal vNewValue As ListObject)
    Set This.Table = vNewValue
End Property

Public Property Get Column() As ListColumn
    Set Column = This.Column
End Property

Public Property Set Column(ByVal vNewValue As ListColumn)
    Set This.Column = vNewValue
End Property

Public Property Get AvailableColumns() As Collection
    Set AvailableColumns = This.AvailableColumns
End Property

Public Property Set AvailableColumns(ByVal vNewValue As Collection)
    Set This.AvailableColumns = vNewValue
End Property

Public Property Get UsedColumns() As Collection
    Set UsedColumns = This.UsedColumns
    End Property
    
Public Property Set UsedColumns(ByVal vNewValue As Collection)
    Set This.UsedColumns = vNewValue
End Property

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
