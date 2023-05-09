Attribute VB_Name = "modTestColumnAnalysis"
'@Folder("Test")
Option Explicit

Public Sub TestColumnAnalysis()
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim ca As ColumnAnalysis
    
    Set lo = GetListObject()
    For Each lc In lo.ListColumns
        Set ca = New ColumnAnalysis
        ca.Analyse lc
        ca.DebugPrint
    Next lc
End Sub
