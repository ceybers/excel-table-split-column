Attribute VB_Name = "modGetSheetNames"
Option Explicit

Public Sub TestGetSheetNames()
    Dim sheetnames as Collection
    Set sheetnames = GetSheetNames(GetListColumn(GetListObject()))
    If not sheetnames is nothing then
        Debug.print "sheetnames count: " & sheetnames.count
        Dim i as long
        For i = 1 to sheetnames.Count
            Debug.print " "; i; " "; sheetnames(i)
            Next
        Else
            Debug.print "sheetnames is nothing"
        End if
End Sub

Public Function GetSheetNames(ByVal lc as ListColumn) as Collection
    Set GetSheetNames = New Collection

    Dim vv as variant
    vv = lc.DataBodyRange.Value2
    Dim i as long
    Dim c As Long
    c = UBound(vv, 1)

    Dim v as variant
    For i = 1 to c
        v = vv(i, 1)
        if VarType(v) = vbString then 
            if IsValidSheetName(v) then
                if not ExistsInCollection(GetSheetNames, v) then
                    GetSheetNames.Add Item:=v, Key:=v
                end if
            end if
        end if
    Next i
End Function

Private Function ExistsInCollection(ByVal coll as Collection, ByVal val as Variant) as Boolean
    Dim v As Variant
    For each v in coll
        if v = val then
            ExistsInCollection = true
            exit function
        end if
    Next v
End Function

Public Sub TestIsValidSheetName()
    TestIsValidSheetNameOne(vbNullString)
    TestIsValidSheetNameOne("history")
    TestIsValidSheetNameOne("thisworksheetnameiswaywaywaytoolong")
    TestIsValidSheetNameOne("\/?*[]:")
    TestIsValidSheetNameOne("Test")
End Sub

Public Sub TestIsValidSheetNameOne(ByVal name as String)
        Debug.Print "IsValidSheetName("; name; ") = "; IsValidSheetName(name)
End Sub

Private Function IsValidSheetName(ByVal name as String) as Boolean
    If name = vbNullString then exit function
    if len(name) > 31 then exit function
    if ucase(name) = "HISTORY" then exit function
    if left$(name, 1) = "'" then exit function

    Dim invalidChar As Variant
    invalidChar = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i as long
    Dim j as long

    For i = 1 to len(name)
        for j = 1 to ubound(invalidChar)
            if mid$(name, i, 1) = invalidChar(j) then exit function
        next j
    Next i

    IsValidSheetName = true
End Function