Attribute VB_Name = "modMain"
'@Folder "Main"
Option Explicit

Public Sub DoSplitTable()
    Dim ViewModel As SplitTableViewModel
    Set ViewModel = New SplitTableViewModel
    ViewModel.Load ActiveWorkbook
    
    ViewModel.DeleteExistingSheets = True
        
    Dim View As frmSplitTable
    Set View = frmSplitTable
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        ProcessViewModel ViewModel
    End If
End Sub
