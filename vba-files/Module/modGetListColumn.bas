Attribute VB_Name = "modGetListColumn"
Option Explicit

Public Sub TestGetListColumn()
    Dim lc As ListColumn
    Set lc = GetListColumn(GetListObject)
    If Not lc Is Nothing Then
        Debug.Print "Lc: " & lc.Name
    Else
        Debug.Print "Lc is nothing"
    End If
End Sub

Public Function GetListColumn(ByVal lo As ListObject) As ListColumn
    Debug.Assert Not lo Is Nothing
    
    Set GetListColumn = lo.listcolumns("Country")
End Function
