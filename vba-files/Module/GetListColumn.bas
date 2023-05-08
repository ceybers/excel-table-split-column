Attribute VB_Name = "modGetListColumn"
Option Explicit

Public Sub TestGetListColumn()
    Dim lc as ListColumn
    Set lc = GetListColumn(GetListObject)
    If not lc is nothing then
        Debug.print "Lc: " & lc.name
    Else
        Debug.print "Lc is nothing"
    End if
End Sub

Public Function GetListColumn(ByVal lo as ListObject) as ListColumn
    Debug.Assert not lo is nothing 
    
    Set GetListColumn = lo.listcolumns("Country")
End Function