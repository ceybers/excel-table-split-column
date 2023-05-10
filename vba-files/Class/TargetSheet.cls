VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Model")
'@PredeclaredId
Option Explicit

Private Type TState
    Name As String
    Used As Boolean
    Suitable As Boolean
End Type
Private This As TState

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
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

Public Property Get Suitable() As Boolean
    Suitable = This.Suitable
End Property

Public Property Let Suitable(ByVal vNewValue As Boolean)
    This.Suitable = vNewValue
End Property
