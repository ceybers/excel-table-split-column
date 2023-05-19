Attribute VB_Name = "modMain"
'@Folder "TableSplit"
Option Explicit

'@EntryPoint "DoSplitTable"
Public Sub DoSplitTable()
    If CheckNoTables(ActiveWorkbook) Then Exit Sub
    If CheckWorkbookProtected(ActiveWorkbook) Then Exit Sub
    
    Dim ViewModel As SplitTableViewModel
    Set ViewModel = New SplitTableViewModel
    ViewModel.Load ActiveWorkbook
    
    TrySelectUserSelectedTable ViewModel
    
    Dim View As frmSplitTable
    Set View = frmSplitTable
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        ProcessViewModel ViewModel
    End If
End Sub

'@Description "Checks if there are no ListObjects in any of the open Excel workbooks."
Private Function CheckNoTables(ByVal Workbook As Workbook) As Boolean
Attribute CheckNoTables.VB_Description = "Checks if there are no ListObjects in any of the open Excel workbooks."
    If ListObjectHelpers.GetAllListObjects(Workbook).Count = 0 Then
        frmNoTables.Show
        CheckNoTables = True
    End If
End Function

'@Description "Checks if a Workbook's structure is protected. If yes, we cannot create new Worksheets, so we display a warning prompt, and gracefully close."
Private Function CheckWorkbookProtected(ByVal Workbook As Workbook) As Boolean
Attribute CheckWorkbookProtected.VB_Description = "Checks if a Workbook's structure is protected. If yes, we cannot create new Worksheets, so we display a warning prompt, and gracefully close."
    If Workbook.ProtectStructure = True Then
        frmWorkbookProtected.Show
        CheckWorkbookProtected = True
    End If
End Function

'@Description "When starting a Split, try and set the initially selected table based on Selection, or then ActiveSheet."
Private Sub TrySelectUserSelectedTable(ByVal ViewModel As SplitTableViewModel)
Attribute TrySelectUserSelectedTable.VB_Description = "When starting a Split, try and set the initially selected table based on Selection, or then ActiveSheet."
    Dim ListObject As ListObject
    Set ListObject = Selection.ListObject
    If ListObject Is Nothing Then
        If ActiveSheet.ListObjects.Count = 0 Then
            Exit Sub
        End If
    End If
    
    Set ListObject = ActiveSheet.ListObjects(1)
    
    ViewModel.TrySelectTableByName ListObject.Name
End Sub
