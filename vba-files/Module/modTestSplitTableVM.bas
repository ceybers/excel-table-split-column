Attribute VB_Name = "modTestSplitTableVM"
'@Folder("Test")
Option Explicit

Public Sub TestSplitTableVM()
    Dim vm As SplitTableViewModel
    Set vm = DummySplitTableViewModel
    
    Dim view As frmSplitTable
    Set view = New frmSplitTable
    
    If view.ShowDialog(vm) Then
        Debug.Print "view.ShowDialog = true"
    Else
        Debug.Print "view.ShowDialog = false"
    End If
End Sub

Private Function DummySplitTableViewModel() As SplitTableViewModel
    Dim Result As SplitTableViewModel
    Set Result = New SplitTableViewModel
    
    Result.Load GetListColumn(GetListObject)
    
    Set DummySplitTableViewModel = Result
End Function
