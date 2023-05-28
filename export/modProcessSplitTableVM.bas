Attribute VB_Name = "modProcessSplitTableVM"
Attribute VB_Description = "Handles some of the optional tasks from a ViewModel before splitting the Table."
'@ModuleDescription "Handles some of the optional tasks from a ViewModel before splitting the Table."
'@Folder "TableSplit.Modules"
Option Explicit

' TODO Check if allowed characters for Tables are the same as Worksheets
Private Const TABLE_PREFIX As String = "tbl"

'@Description "Applies the instructions in the ViewModel to the Workbook."
Public Sub ProcessSplitTableVM(ByVal ViewModel As SplitTableViewModel)
Attribute ProcessSplitTableVM.VB_Description = "Applies the instructions in the ViewModel to the Workbook."
    Dim mCalculationMode As Long
    Dim mScreenUpdating As Boolean
    mCalculationMode = Application.Calculation
    mScreenUpdating = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    TryRemoveOtherSheets ViewModel
    TryRemoveExistingSheets ViewModel
    
    SplitTable ViewModel
    
    Application.Calculation = mCalculationMode
    Application.ScreenUpdating = mScreenUpdating
End Sub

'@Description "Optionally removes all worksheets in the workbook except the one we are splitting."
Private Sub TryRemoveOtherSheets(ByVal ViewModel As SplitTableViewModel)
Attribute TryRemoveOtherSheets.VB_Description = "Optionally removes all worksheets in the workbook except the one we are splitting."
    Log.Message "TryRemoveOtherSheets"
    
    If ViewModel.RemoveOtherSheets = False Then Exit Sub
    
    Dim ListObject As ListObject
    Set ListObject = ViewModel.AvailableTables.Selected
    
    Dim KeepWorksheet As Worksheet
    Set KeepWorksheet = ListObject.Parent
    
    Dim Worksheet As Worksheet
    For Each Worksheet In KeepWorksheet.Parent.Worksheets
        If Worksheet.Name <> KeepWorksheet.Name Then
            Application.DisplayAlerts = False
            Worksheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Worksheet
End Sub

'@Description "Optionally removes existing worksheets with the same name as the ones we are creating."
Private Sub TryRemoveExistingSheets(ByVal ViewModel As SplitTableViewModel)
Attribute TryRemoveExistingSheets.VB_Description = "Optionally removes existing worksheets with the same name as the ones we are creating."
    Log.Message "TryRemoveExistingSheets"
    If ViewModel.DeleteExistingSheets = False Then Exit Sub
    
    Dim ListObject As ListObject
    Dim SheetNames As Collection

    Set ListObject = ViewModel.AvailableTables.Selected
    Set SheetNames = ViewModel.TargetSheets.GetSelectedSheetNames

    Dim Worksheet As Worksheet
    For Each Worksheet In ListObject.Parent.Parent.Worksheets
        If ExistsInCollection(SheetNames, Worksheet.Name) Then
            Application.DisplayAlerts = False
            Worksheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Worksheet
End Sub

'@Description "Loops through the unique values one by one, creating new Reduced worksheets, and inserting them in the correct order."
Private Sub SplitTable(ByVal ViewModel As SplitTableViewModel)
Attribute SplitTable.VB_Description = "Loops through the unique values one by one, creating new Reduced worksheets, and inserting them in the correct order."
    Log.Message "SplitTable"
    Dim ListColumn As ListColumn
    Dim SheetNames As Collection

    Set ListColumn = ViewModel.AvailableColumns.Selected
    Set SheetNames = ViewModel.TargetSheets.GetSelectedSheetNames

    Dim ListObject As ListObject
    Set ListObject = ListColumn.Parent
    
    Dim Workbook As Workbook
    Set Workbook = ListObject.Parent.Parent
    
    Dim SourceWorksheet As Worksheet
    Set SourceWorksheet = ListObject.Parent

    Dim PreviousWorksheet As Worksheet
    Set PreviousWorksheet = SourceWorksheet

    Dim SheetsToSplit As Long
    SheetsToSplit = SheetNames.Count
    
    Dim CurrentSheetNumber As Long
    Dim SheetsActuallyTransfered As Long
    
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
            ReduceWorksheet NewWorksheet, ListColumn.Name, SheetName
            Set PreviousWorksheet = NewWorksheet
            SheetsActuallyTransfered = SheetsActuallyTransfered + 1
        End If
    Next SheetName
    
    Log.Message "CurrentSheetNumber: " & CurrentSheetNumber, "SplitTable", Verbose_Level
    Log.Message "SheetsActuallyTransfered: " & SheetsActuallyTransfered, "SplitTable", Verbose_Level
    
    Dim SheetsNotTransferred As Long
    SheetsNotTransferred = CurrentSheetNumber - SheetsActuallyTransfered
    
    If SheetsNotTransferred > 0 Then
        Dim NotTransferredMessage As String
        NotTransferredMessage = SheetsNotTransferred & " sheet(s) were not replaced!"
        MsgBox NotTransferredMessage, vbExclamation + vbOKOnly, "Table Split Tool"
    End If
    
    If SheetsActuallyTransfered > 0 Then
        ProgressBarDialog.UpdateProgress 1#
        ProgressBarDialog.Show vbModal
    End If
    
    Application.StatusBar = vbNullString
    SourceWorksheet.Activate
End Sub

'@Description "Filteres a Worksheet by a given ListColumn to a specific Criteria, then removes all the unfiltered rows."
Private Sub ReduceWorksheet(ByVal Worksheet As Worksheet, ByVal ListColumnName As String, ByVal SheetName As String)
Attribute ReduceWorksheet.VB_Description = "Filteres a Worksheet by a given ListColumn to a specific Criteria, then removes all the unfiltered rows."
    Log.Message "ReduceWorksheet"
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
        Application.Intersect(RangeToRemove.EntireRow, ListObject.DataBodyRange).Delete
    End If
    Application.DisplayAlerts = True

    ListObject.Range.Autofilter Field:=ListColumnIndex
    
    ListObject.ListColumns.Item(ListColumnName).DataBodyRange.Select
End Sub
