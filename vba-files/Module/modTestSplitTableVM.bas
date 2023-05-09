Attribute VB_Name = "modTestSplitTableVM"
'@Folder("Test")
Option Explicit

Public Sub TestSplitTableVM()
    
End Sub

Private Function DummySplitTableViewModel() As SplitTableViewModel
    Dim Result As SplitTableViewModel
    Set Result = New SplitTableViewModel
    
    Result.Load GetListColumn(GetListObject)
    
    Set DummySplitTableViewModel = Result
End Function
