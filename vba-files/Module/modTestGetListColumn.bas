Attribute VB_Name = "modTestGetListColumn"
'@Folder "Test"
Option Explicit

'@Description "NYI"
Public Function GetListColumn(ByVal lo As ListObject) As ListColumn
Attribute GetListColumn.VB_Description = "NYI"
    Debug.Assert Not lo Is Nothing
    
    ' TODO Implement this once UI picker is done
    Set GetListColumn = lo.ListColumns("Country")
    
    Debug.Assert False
End Function
