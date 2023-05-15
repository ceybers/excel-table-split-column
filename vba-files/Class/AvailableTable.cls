VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Model")
'@PredeclaredId
Option Explicit

Private Type TState
    'ListObject As ListObject
    Name As String
    Suitable As Boolean
End Type
Private This As TState

'Public Property Get ListObject() As ListObject
'    Set ListObject = This.ListObject
'End Property

'Public Property Set ListObject(ByVal vNewValue As ListObject)
'    Set This.ListObject = vNewValue
'End Property

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Get Suitable() As Boolean
    Suitable = This.Suitable
End Property

Public Property Let Suitable(ByVal vNewValue As Boolean)
    This.Suitable = vNewValue
End Property

Public Function Create(ByVal ListObject As ListObject) As AvailableTable
    Set Create = New AvailableTable
    With Create
        'Set .ListObject = ListObject
        .Name = ListObject.Name & " (" & ListObject.Parent.Name & ")"
        .Suitable = IsSuitable(ListObject)
    End With
End Function

'@Description "Does this table have more than one row AND protection is disabled?"
Private Function IsSuitable(ByVal ListObject As ListObject) As Boolean
Attribute IsSuitable.VB_Description = "Does this table have more than one row AND protection is disabled?"
    'Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    
    Set Worksheet = ListObject.Parent
    'Set Workbook = Worksheet.Parent
    
    If Worksheet.ProtectContents = True Then Exit Function
    'If Workbook.ProtectStructure = True Then Exit Function
    If ListObject.ListRows.Count = 1 Then Exit Function
    
    IsSuitable = True
End Function
