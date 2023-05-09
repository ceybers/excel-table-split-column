Attribute VB_Name = "modMain"
'@Folder "Main"
Option Explicit

Public Sub SplitTable()
    Dim lo As ListObject
    Set lo = GetListObject

    Dim lc As ListColumn
    Set lc = lo.ListColumns("Country")

    DoSplitTable lo, lc
End Sub

