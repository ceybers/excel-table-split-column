Attribute VB_Name = "modGetListColumn"
Option Explicit

Public Function GetListColumn(ByVal lo As ListObject) As ListColumn
    Debug.Assert Not lo Is Nothing
    
    ' TODO Implement this once UI picker is done
    Set GetListColumn = lo.ListColumns("Country")
End Function
