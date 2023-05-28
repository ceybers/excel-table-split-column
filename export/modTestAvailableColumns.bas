Attribute VB_Name = "modTestAvailableColumns"
'@Folder "Test"
Option Explicit

Public Sub TestAvailableColumns()
    Dim ViewModel As SplitTableViewModel
    Set ViewModel = New SplitTableViewModel
    ViewModel.Load ActiveWorkbook
    
    Dim View As frmTestAvailableColumns
    Set View = frmTestAvailableColumns
    
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
    Debug.Print "NYI"
    ' TODO Only split out selected items
    'DoSplitTable ViewModel.SelectedListColumn.Parent, ViewModel.SelectedListColumn
End Sub

