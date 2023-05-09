Attribute VB_Name = "modGetSheetNames"
Option Explicit

'@Description "Creates a collection of unique and valid sheet names from a ListColumn's DataBodyRange."
Public Function GetSheetNames(ByVal lc As ListColumn) As Collection
    Set GetSheetNames = New Collection

    Dim vv As Variant
    vv = lc.DataBodyRange.Value2
    Dim i As Long
    Dim c As Long
    c = UBound(vv, 1)

    Dim v As Variant
    For i = 1 To c
        v = vv(i, 1)
        If VarType(v) = vbString Then
            If IsValidSheetName(v) Then
                If Not ExistsInCollection(GetSheetNames, v) Then
                    GetSheetNames.Add Item:=v, Key:=v
                End If
            End If
        End If
    Next i
End Function
