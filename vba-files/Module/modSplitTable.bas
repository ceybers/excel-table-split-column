Attribute VB_Name = "modSplitTable"
Attribute VB_Description = "Given a ListColumn and a Collection of worksheet Names, does the actual splitting."
'@ModuleDescription "Given a ListColumn and a Collection of worksheet Names, does the actual splitting."
'@Folder "TableSplit.Modules"
Option Explicit

Private Const TABLE_PREFIX As String = "tbl"

Public Sub SplitTable(ByVal ListColumn As ListColumn, ByVal SheetNames As Collection)
    Dim ListObject As ListObject
    Set ListObject = ListColumn.Parent
    
    Dim Workbook As Workbook
    Set Workbook = ListObject.Parent.Parent
    
    Dim SourceWorksheet As Worksheet
    Set SourceWorksheet = ListObject.Parent

    Dim PreviousWorksheet As Worksheet
    Set PreviousWorksheet = SourceWorksheet

    Dim SheetsToSplit As Long
    Dim CurrentSheetNumber As Long
    SheetsToSplit = SheetNames.Count
    
    Dim ProgressBarDialog As frmProgress
    Set ProgressBarDialog = New frmProgress
    'ProgressBarDialog.Show vbModeless

    Dim NewWorksheet As Worksheet
    Dim SheetName As Variant
    For Each SheetName In SheetNames
    CurrentSheetNumber = CurrentSheetNumber + 1
        ProgressBarDialog.UpdateProgress CDbl(CurrentSheetNumber / SheetsToSplit)
        Application.StatusBar = "Progress: " & CurrentSheetNumber & " of " & SheetsToSplit
        DoEvents
        
        If Not SheetExists(SourceWorksheet.Parent, SheetName) Then
            SourceWorksheet.Copy After:=PreviousWorksheet
            Set NewWorksheet = Workbook.Worksheets.Item(PreviousWorksheet.Index + 1)
            NewWorksheet.Name = SheetName
            FilterWorksheet NewWorksheet, ListColumn.Name, SheetName
            Set PreviousWorksheet = NewWorksheet
        End If
    Next SheetName
    
    ProgressBarDialog.UpdateProgress 1#
    ProgressBarDialog.Show vbModal
    
    SourceWorksheet.Activate
End Sub

Private Sub FilterWorksheet(ByVal Worksheet As Worksheet, ByVal ListColumnName As String, ByVal SheetName As String)
    Dim ListObject As ListObject
    Set ListObject = Worksheet.ListObjects.Item(1) ' TODO

    ListObject.Name = TABLE_PREFIX & SheetName ' TODO Move this out of this procedure
    
    Dim ListColumnIndex As Long
    ListColumnIndex = ListObject.ListColumns.Item(ListColumnName).Index

    ListObject.Range.Autofilter Field:=ListColumnIndex, Criteria1:="<>" & SheetName, Operator:=xlOr

    Dim RangeToRemove As Range
    Set RangeToRemove = ListObject.ListColumns.Item(ListColumnName).DataBodyRange.SpecialCells(xlCellTypeVisible)
    Application.DisplayAlerts = False
    If Not RangeToRemove Is Nothing Then
        RangeToRemove.Rows.Delete
    End If
    Application.DisplayAlerts = True

    ListObject.Range.Autofilter Field:=ListColumnIndex
End Sub
