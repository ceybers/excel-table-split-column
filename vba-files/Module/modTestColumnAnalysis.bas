Attribute VB_Name = "modTestColumnAnalysis"
'@Folder("Test")
Option Explicit

Public Sub TestColumnAnalysis()
    Dim lo As ListObject
    Dim lc As ListColumn
    
    Dim vType As Long
    Dim Uniqueness As Double
    
    Set lo = GetListObject()
    For Each lc In lo.ListColumns
        vType = AnalyseColumnVarType(lc)
        Uniqueness = AnalyseColumnUniqueness(lc, vType)
        Debug.Print " "; Uniqueness
    Next lc
End Sub

Private Function AnalyseColumnVarType(ByVal lc As ListColumn) As Long
    Dim colVarType As Long
    Dim thisVarType As Long
    
    Debug.Print lc.Name
    Dim vv As Variant
    vv = lc.DataBodyRange.Value2
    
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

Private Function AnalyseColumnUniqueness(ByVal lc As ListColumn, ByVal vType As Long) As Double
    Dim AllValues As Collection
    Dim UniqueValues As Collection
    
    Set AllValues = New Collection
    Set UniqueValues = New Collection
    
    Dim v As Variant
    Dim vv As Variant
    vv = lc.DataBodyRange.Value2
    
    For Each v In vv
        If VarType(v) = (vType - vbArray) Then
            If Not ExistsInCollection(AllValues, v) Then
                UniqueValues.Add Item:=v
            End If
            AllValues.Add Item:=v
        End If
    Next v
    
    If AllValues.Count = 0 Then Exit Function
    
    AnalyseColumnUniqueness = UniqueValues.Count / AllValues.Count
End Function
