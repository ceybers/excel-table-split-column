VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model"
Option Explicit

Private Type TState
    Name As String
    ColumnType As Long
    Uniqueness As Long
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Get ColumnType() As Long
    ColumnType = This.ColumnType
End Property

Public Property Let ColumnType(ByVal vNewValue As Long)
    This.ColumnType = vNewValue
End Property

Public Property Get Uniqueness() As Long
    Uniqueness = This.Uniqueness
End Property

Public Property Let Uniqueness(ByVal vNewValue As Long)
    This.Uniqueness = vNewValue
End Property
      
