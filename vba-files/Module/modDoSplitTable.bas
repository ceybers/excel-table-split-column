Attribute VB_Name = "modDoSplitTable"
'@Folder "SplitTable"
Option Explicit

Private Const TABLE_PREFIX As String = "tbl"

Public Sub DoSplitTable(ByVal ListObject As ListObject, ByVal ListColumn As ListColumn)
    ' TODO ListObject is redundant, derive it from ListColumn
    ' TODO Remove implicit references to ActiveWorkbook and ActiveSheet
    Dim sheetnames As Collection
    Set sheetnames = GetSheetNames(ListColumn)
    
    Dim SourceWorksheet As Worksheet
    Set SourceWorksheet = ListObject.Parent

    Dim PreviousWorksheet As Worksheet
    Set PreviousWorksheet = SourceWorksheet

    Dim NewWorksheet As Worksheet
    Dim SheetName As Variant
    For Each SheetName In sheetnames
        TryRemoveSheet ActiveWorkbook, SheetName
        SourceWorksheet.Copy After:=PreviousWorksheet
        Set NewWorksheet = Worksheets.Item(PreviousWorksheet.Index + 1)
        NewWorksheet.Name = SheetName
        FilterWorksheet NewWorksheet, ListColumn.Name, SheetName
        Set PreviousWorksheet = NewWorksheet
    Next SheetName

    SourceWorksheet.Activate
End Sub

Private Sub FilterWorksheet(ByVal Worksheet As Worksheet, ByVal ListColumnName As String, ByVal SheetName As String)
    Dim ListObject As ListObject
    Set ListObject = Worksheet.ListObjects(1) ' TODO

    ListObject.Name = TABLE_PREFIX & SheetName ' TODO Move this out of this procedure
    
    Dim ListColumnIndex As Long
    ListColumnIndex = ListObject.ListColumns(ListColumnName).Index

    ListObject.Range.Autofilter Field:=ListColumnIndex, Criteria1:="<>" & SheetName, Operator:=xlOr

    Dim RangeToRemove As Range
    Set RangeToRemove = ListObject.DataBodyRange.SpecialCells(xlCellTypeVisible)
    Application.DisplayAlerts = False
    If Not RangeToRemove Is Nothing Then RangeToRemove.Rows.Delete
    Application.DisplayAlerts = True

    ListObject.Range.Autofilter Field:=ListColumnIndex
End Sub
