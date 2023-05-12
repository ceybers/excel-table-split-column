VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Model")
Option Explicit

Private Type TState
    ListColumn As ListColumn
    Name As String
    Value2 As Variant
    VarType As Long
    AllValues As Collection
    UniqueValues As Collection
End Type
Private This As TState

Public Property Get ListColumn() As ListColumn
    Debug.Assert Not This.ListColumn Is Nothing
    Set ListColumn = This.ListColumn
End Property

Public Property Get Uniqueness() As Double
    If This.AllValues.Count = 0 Then Exit Property
    Uniqueness = This.UniqueValues.Count / This.AllValues.Count
End Property

Public Property Get UniqueCount() As Long
    UniqueCount = This.UniqueValues.Count
End Property

Public Property Get AllCount() As Long
    AllCount = This.AllValues.Count
End Property

Public Property Get ColumnVarType() As Long
    ColumnVarType = This.VarType
End Property

Public Sub Analyse(ByVal lc As ListColumn)
    Set This.ListColumn = lc
    This.Value2 = This.ListColumn.DataBodyRange.Value2
    This.VarType = AnalyseColumnVarType(This.Value2)
    AnalyseColumnUniqueness
End Sub

Private Function AnalyseColumnVarType(ByVal vv As Variant) As Long
    Dim colVarType As Long
    Dim thisVarType As Long
  
    Dim i As Long
    Dim c As Long
    c = UBound(vv, 1)
    
    Dim v As Variant
    For i = 1 To c
        v = vv(i, 1)
        thisVarType = VarType(v)
        
        If thisVarType > vbEmpty Then 'Exclude vbEmpty from resetting other VarType
            If colVarType = vbEmpty Then
                colVarType = thisVarType
            Else
                If colVarType <> thisVarType Then
                    colVarType = vbVariant
                End If
            End If
        End If
    Next i
    
    colVarType = colVarType + vbArray
    AnalyseColumnVarType = colVarType
End Function

Private Sub AnalyseColumnUniqueness()
    Debug.Assert This.AllValues Is Nothing
    Debug.Assert This.UniqueValues Is Nothing
    
    Set This.AllValues = New Collection
    Set This.UniqueValues = New Collection
    
    Dim v As Variant
    For Each v In This.Value2
        If VarType(v) = (This.VarType - vbArray) Then
            If Not ExistsInCollection(This.AllValues, v) Then
                This.UniqueValues.Add Item:=v
            End If
            This.AllValues.Add Item:=v
        End If
    Next v
End Sub

Public Sub DebugPrint()
    Debug.Print "ListColumn.Name: "; This.ListColumn.Name
    Debug.Print " UniqueCount: "; UniqueCount
    Debug.Print " AllCount: "; AllCount
    Debug.Print " Uniqueness: "; FormatPercent(Uniqueness, 0)
    Debug.Print " ColumnVarType: "; VarTypeToString(ColumnVarType)
End Sub

Public Function GetTargetSheets() As Collection
    Set GetTargetSheets = New Collection
    
    Dim TargetSheetName As Variant
    For Each TargetSheetName In This.UniqueValues
        Dim TargetSheet As TargetSheet
        Set TargetSheet = New TargetSheet
        TargetSheet.Name = TargetSheetName
        TargetSheet.Suitable = WorksheetHelpers.IsValidSheetName(TargetSheetName)
        TargetSheet.Used = TargetSheet.Suitable
        GetTargetSheets.Add Item:=TargetSheet, Key:=TargetSheetName
    Next TargetSheetName
End Function

Public Function Create(ByVal ListColumn As ListColumn) As ColumnAnalysis
    Set Create = New ColumnAnalysis
    Create.Analyse ListColumn
End Function
