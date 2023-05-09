Attribute VB_Name = "WorksheetHelpers"
'@Folder "Helpers"
Option Explicit

Public Function IsValidSheetName(ByVal Name As String) As Boolean
    If Name = "" Then Exit Function
    If Len(Name) > 31 Then Exit Function
    If UCase(Name) = "HISTORY" Then Exit Function
    If Left$(Name, 1) = "'" Then Exit Function

    Dim invalidChar As Variant
    invalidChar = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long
    Dim j As Long

    For i = 1 To Len(Name)
        For j = 1 To UBound(invalidChar)
            If Mid$(Name, i, 1) = invalidChar(j) Then Exit Function
        Next j
    Next i

    IsValidSheetName = True
End Function

