Attribute VB_Name = "modGetListObject"
Option Explicit

Public Sub TestGetListObject()
    Dim lo as ListObject
    Set lo = GetListObject()
    If not lo is nothing then
        Debug.print "Lo: " & lo.name
    Else
        Debug.print "Lo is nothing"
    End if
End Sub

Public Function GetListObject() as ListObject
    Set GetListObject = TryGetSelectedListObject()
    If Not GetListObject is Nothing Then Exit Function

    Set GetListObject = TryGetListObjectOnSheet()
    If Not GetListObject is Nothing Then Exit Function

    Set GetListObject = TryGetOnlyListObjectInWorkbook()
    If Not GetListObject is Nothing Then Exit Function
End Function

Public Function TryGetSelectedListObject() as ListObject
    Set TryGetSelectedListObject = Selection.ListObject
End Function

Public Function TryGetListObjectOnSheet() as ListObject
    If ActiveSheet.ListObjects.Count = 1 then
        Set TryGetListObjectOnSheet = Activesheet.ListObjects(1)
    End If  
End Function

Public Function TryGetOnlyListObjectInWorkbook() as ListObject
    Dim result as ListObject

    Dim ws as Worksheet
    For each ws in Activeworkbook.Worksheets
        if ws.ListObjects.Count = 1 then
            ' Exit without returning a result if there is already a ListObject set
            if not result is nothing then exit function

            set result = ws.listobjects(1)
        end if
    Next ws

    Set TryGetOnlyListObjectInWorkbook = result
End Function