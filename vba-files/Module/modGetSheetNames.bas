Attribute VB_Name = "modGetSheetNames"
'@Folder "TableSplit.Modules"
Option Explicit

'@Description "Creates a collection of unique and valid sheet names from a ListColumn's DataBodyRange."
Public Function GetSheetNames(ByVal ListColumn As ListColumn) As Collection
Attribute GetSheetNames.VB_Description = "Creates a collection of unique and valid sheet names from a ListColumn's DataBodyRange."
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
