Attribute VB_Name = "ListColumnHelpers"
'@Folder "TableSplit.Helpers"
Option Explicit

'@Description "Returns a collection of unique values in a ListColumn's DataBodyRange. Filtered to only include values that are valid Sheet Names."
Public Function GetSheetNames(ByVal ListColumn As ListColumn) As Collection
Attribute GetSheetNames.VB_Description = "Returns a collection of unique values in a ListColumn's DataBodyRange. Filtered to only include values that are valid Sheet Names."
    Set GetSheetNames = New Collection

    Dim Value2 As Variant
    Value2 = ListColumn.DataBodyRange.Value2
    
    Dim i As Long
    Dim ThisValue2 As Variant
    For i = 1 To UBound(Value2, 1)
        ThisValue2 = Value2(i, 1)
        If VarType(ThisValue2) = vbString Then
            If IsValidSheetName(ThisValue2) Then
                If Not ExistsInCollection(GetSheetNames, ThisValue2) Then
                    GetSheetNames.Add Item:=ThisValue2, Key:=ThisValue2
                End If
            End If
        End If
    Next i
End Function
