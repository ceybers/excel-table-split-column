Attribute VB_Name = "modTestModules"
'@Folder "Test"
Option Explicit

Public Sub TestColumnAnalysis()
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim ca As ColumnAnalysis
    
    Set lo = GetListObject()
    For Each lc In lo.ListColumns
        Set ca = New ColumnAnalysis
        ca.Analyse lc
        ca.DebugPrint
    Next lc
End Sub

Public Sub TestDoSplitTable()
    Dim lo As ListObject
    Set lo = GetListObject

    Dim lc As ListColumn
    Set lc = GetListColumn(lo)

    DoSplitTable lo, lc
End Sub

Public Sub TestGetListColumn()
    Dim lc As ListColumn
    Set lc = GetListColumn(GetListObject)
    If Not lc Is Nothing Then
        Debug.Print "Lc: " & lc.Name
    Else
        Debug.Print "Lc is nothing"
    End If
End Sub

Public Sub TestGetListObject()
    Dim lo As ListObject
    Set lo = GetListObject()
    If Not lo Is Nothing Then
        Debug.Print "Lo: " & lo.Name
    Else
        Debug.Print "Lo is nothing"
    End If
End Sub

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

Public Sub TestIsValidSheetName()
    TestIsValidSheetNameOne vbNullString
    TestIsValidSheetNameOne "history"
    TestIsValidSheetNameOne "thisworksheetnameiswaywaywaytoolong"
    TestIsValidSheetNameOne "\/?*[]:"
    TestIsValidSheetNameOne "Test"
End Sub

Private Sub TestIsValidSheetNameOne(ByVal Name As String)
        Debug.Print "IsValidSheetName("; Name; ") = "; IsValidSheetName(Name)
End Sub
