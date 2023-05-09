Attribute VB_Name = "modGetSheetNames"
Option Explicit

Public Sub TestGetSheetNames()
    Dim sheetnames As Collection
    Set sheetnames = GetSheetNames(GetListColumn(GetListObject()))
    If Not sheetnames Is Nothing Then
        Debug.Print "sheetnames count: " & sheetnames.Count
        Dim i As Long
        For i = 1 To sheetnames.Count
            Debug.Print " "; i; " "; sheetnames(i)
            Next
        Else
            Debug.Print "sheetnames is nothing"
        End If
End Sub

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

Private Function ExistsInCollection(ByVal coll As Collection, ByVal val As Variant) As Boolean
    Dim v As Variant
    For Each v In coll
        If v = val Then
            ExistsInCollection = True
            Exit Function
        End If
    Next v
End Function

Public Sub TestIsValidSheetName()
    TestIsValidSheetNameOne ("")
    TestIsValidSheetNameOne ("history")
    TestIsValidSheetNameOne ("thisworksheetnameiswaywaywaytoolong")
    TestIsValidSheetNameOne ("\/?*[]:")
    TestIsValidSheetNameOne ("Test")
End Sub

Public Sub TestIsValidSheetNameOne(ByVal Name As String)
        Debug.Print "IsValidSheetName("; Name; ") = "; IsValidSheetName(Name)
End Sub

Private Function IsValidSheetName(ByVal Name As String) As Boolean
    If Name = "" Then Exit Function
    If Len(Name) > 31 Then Exit Function
    If ucase(Name) = "HISTORY" Then Exit Function
    If left$(Name, 1) = "'" Then Exit Function

    Dim invalidChar As Variant
    invalidChar = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long
    Dim j As Long

    For i = 1 To Len(Name)
        For j = 1 To UBound(invalidChar)
            If mid$(Name, i, 1) = invalidChar(j) Then Exit Function
        Next j
    Next i

    IsValidSheetName = True
End Function
