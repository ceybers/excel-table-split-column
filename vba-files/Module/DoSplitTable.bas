Attribute VB_Name = "modDoSplitTable"
Option Explicit

Public Sub TestDoSplitTable()
    Dim lo as ListObject
    set lo = GetListObject

    Dim lc as ListColumn
    set lc = GetListColumn(lo)

    DoSplitTable lo, lc
End Sub

Public Sub DoSplitTable(ByVal lo as ListObject, ByVal lc as ListColumn)
    Dim sheetNames as Collection
    Set sheetNames = GetSheetNames(lc)
    
    Dim sourceWorksheet as Worksheet
    Set sourceWorksheet = lo.Parent

    Dim previousWorksheet as worksheet
    Set previousWorksheet = sourceWorksheet

    Dim newWorksheet as Worksheet
    Dim v as variant
    For each v in sheetNames
        TryRemoveSheet v
        sourceWorksheet.Copy After:=previousWorksheet
        Set newWorksheet = Worksheets(previousWorksheet.Index + 1)
        'set newWorksheet = sourceWorksheet.Copy(,previousWorksheet)
        newWorksheet.name = v
        FilterWorksheet newWorksheet, lc.Name, v
        Set previousWorksheet = newWorksheet
    Next v

    sourceWorksheet.activate
End Sub


Private Function TryRemoveSheet(ByVal name as String) as Boolean
    Dim ws as Worksheet
    For each ws in ActiveWorkbook.Worksheets
        if ws.name = name then
            application.displayalerts = false
            ws.Delete
            application.displayalerts = True
            TryRemoveSheet = True
            Exit Function
        end if
    Next ws
End Function

Private Sub FilterWorksheet(ByVal ws as Worksheet, ByVal lcName as String, ByVal val as String)
    Dim lo as ListObject
    Set lo = ws.ListObjects(1)

    Dim lcIndex as Long
    lcIndex = lo.ListColumns(lcName).Index

    lo.name = "tbl" & Val
    lo.Range.Autofilter Field:=lcIndex, Criteria1:="<>"&val, Operator:= xlOr

    application.DisplayAlerts = false
    lo.Databodyrange.SpecialCells(xlCellTypeVisible).Rows.Delete
    application.DisplayAlerts = true

    lo.Range.Autofilter Field:=lcIndex
End Sub