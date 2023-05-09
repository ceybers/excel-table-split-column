Attribute VB_Name = "modDoSplitTable"
Option Explicit

Public Sub DoSplitTable(ByVal lo As ListObject, ByVal lc As ListColumn)
    Dim sheetnames As Collection
    Set sheetnames = GetSheetNames(lc)
    
    Dim sourceWorksheet As Worksheet
    Set sourceWorksheet = lo.Parent

    Dim previousWorksheet As Worksheet
    Set previousWorksheet = sourceWorksheet

    Dim newWorksheet As Worksheet
    Dim v As Variant
    For Each v In sheetnames
        TryRemoveSheet v
        sourceWorksheet.Copy After:=previousWorksheet
        Set newWorksheet = Worksheets(previousWorksheet.Index + 1)
        'set newWorksheet = sourceWorksheet.Copy(,previousWorksheet)
        newWorksheet.Name = v
        FilterWorksheet newWorksheet, lc.Name, v
        Set previousWorksheet = newWorksheet
    Next v

    sourceWorksheet.Activate
End Sub

Private Function TryRemoveSheet(ByVal Name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Activeworkbook.Worksheets
        If ws.Name = Name Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            TryRemoveSheet = True
            Exit Function
        End If
    Next ws
End Function

Private Sub FilterWorksheet(ByVal ws As Worksheet, ByVal lcName As String, ByVal val As String)
    Dim lo As ListObject
    Set lo = ws.listobjects(1)

    Dim lcIndex As Long
    lcIndex = lo.ListColumns(lcName).Index

    lo.Name = "tbl" & val
    lo.Range.Autofilter Field:=lcIndex, Criteria1:="<>" & val, Operator:=xlOr

    Dim delRange As Range
    Set delRange = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
    Application.DisplayAlerts = False
    If Not delRange Is Nothing Then delRange.Rows.Delete
    Application.DisplayAlerts = True

    lo.Range.Autofilter Field:=lcIndex
End Sub
