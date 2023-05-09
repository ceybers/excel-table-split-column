VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model")
Option Explicit

Private Type TState
    Name As String
    Used As Boolean
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Get Used() As Boolean
    Used = This.Used
End Property

Public Property Let Used(ByVal vNewValue As Boolean)
    This.Used = vNewValue
End Property
