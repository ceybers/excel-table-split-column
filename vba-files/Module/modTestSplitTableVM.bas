Attribute VB_Name = "modTestSplitTableVM"
'@Folder("Test")
Option Explicit

Public Sub TestSplitTableVM()
    Dim ViewModel As SplitTableViewModel
    Set ViewModel = New SplitTableViewModel
    ViewModel.Load ActiveWorkbook
    
    Dim View As frmSplitTable
    Set View = frmSplitTable
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        Debug.Print "ViewAsInterface.ShowDialog(ViewModel) = True"
        ProcessViewModel ViewModel
    Else
        Debug.Print "ViewAsInterface.ShowDialog(ViewModel) = False"
    End If
End Sub

Private Sub ProcessViewModel(ByVal ViewModel As SplitTableViewModel)
    ' TODO Only split out selected items
    DoSplitTable ViewModel.SelectedListColumn.Parent, ViewModel.SelectedListColumn
End Sub
