Attribute VB_Name = "modMain"
'@Folder "Main"
Option Explicit

'@EntryPoint "DoSplitTable"
Public Sub DoSplitTable()
    If CheckNoTables(ActiveWorkbook) Then Exit Sub
    If CheckWorkbookProtected(ActiveWorkbook) Then Exit Sub
    
    Dim ViewModel As SplitTableViewModel
    Set ViewModel = New SplitTableViewModel
    ViewModel.Load ActiveWorkbook
        
    Dim View As frmSplitTable
    Set View = frmSplitTable
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        ProcessViewModel ViewModel
    End If
End Sub

Private Function CheckNoTables(ByVal Workbook As Workbook) As Boolean
    If ListObjectHelpers.GetAllListObjects(Workbook).Count = 0 Then
        frmNoTables.Show
        CheckNoTables = True
    End If
End Function

Private Function CheckWorkbookProtected(ByVal Workbook As Workbook) As Boolean
    If Workbook.ProtectStructure = True Then
        frmWorkbookProtected.Show
        CheckWorkbookProtected = True
    End If
End Function
