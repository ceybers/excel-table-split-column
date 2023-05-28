VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "TableSplit.Model"
'@PredeclaredId
Option Explicit

Private Type TState
    ListColumn As ListColumn
    Name As String
    UniqueValueCount As Long
    
    Selected As Boolean
    Hidden As Boolean
    Suitable As Boolean
    Index As Long
End Type
Private This As TState

Public Property Get ListColumn() As ListColumn
    Set ListColumn = This.ListColumn
End Property

Public Property Set ListColumn(ByVal vNewValue As ListColumn)
    Set This.ListColumn = vNewValue
End Property

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Get UniqueValueCount() As Long
    UniqueValueCount = This.UniqueValueCount
End Property

Public Property Let UniqueValueCount(ByVal vNewValue As Long)
    This.UniqueValueCount = vNewValue
End Property

Public Property Get Hidden() As Boolean
    Hidden = This.Hidden
End Property

Public Property Let Hidden(ByVal vNewValue As Boolean)
    This.Hidden = vNewValue
End Property

Public Property Get Selected() As Boolean
    Selected = This.Selected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    This.Selected = vNewValue
End Property

Public Property Get Suitable() As Boolean
    Suitable = This.Suitable
End Property

Public Property Let Suitable(ByVal vNewValue As Boolean)
    This.Suitable = vNewValue
End Property

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Let Index(ByVal vNewValue As Long)
    This.Index = vNewValue
End Property

Private Function IsSuitable(ByVal ColumnAnalysis As ColumnAnalysis) As Boolean
    If ColumnAnalysis.AllCount < 1 Then Exit Function
    If ColumnAnalysis.UniqueCount > 10 Then Exit Function
    If ColumnAnalysis.ColumnVarType <> (vbArray + vbString) Then Exit Function
    
    IsSuitable = True
End Function

Public Function Create(ByVal ListColumn As ListColumn) As AvailableColumn
    Dim ThisColumnAnalysis As ColumnAnalysis
    Set ThisColumnAnalysis = ColumnAnalysis.Create(ListColumn)
    
    Set Create = New AvailableColumn
    With Create
        Set .ListColumn = ListColumn
        .Name = ListColumn.Name
        .UniqueValueCount = ThisColumnAnalysis.UniqueCount
        .Hidden = ListColumn.Range.EntireColumn.Hidden
        .Selected = False
        .Suitable = IsSuitable(ThisColumnAnalysis)
    End With
End Function
