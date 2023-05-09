Attribute VB_Name = "modGetListObject"
Option Explicit

Public Sub TestGetListObject()
    Dim lo As ListObject
    Set lo = GetListObject()
    If Not lo Is Nothing Then
        Debug.Print "Lo: " & lo.Name
    Else
        Debug.Print "Lo is nothing"
    End If
End Sub

Public Function GetListObject() As ListObject
    Set GetListObject = TryGetSelectedListObject()
    If Not GetListObject Is Nothing Then Exit Function

    Set GetListObject = TryGetListObjectOnSheet()
    If Not GetListObject Is Nothing Then Exit Function

    Set GetListObject = TryGetOnlyListObjectInWorkbook()
    If Not GetListObject Is Nothing Then Exit Function
End Function

Public Function TryGetSelectedListObject() As ListObject
    Set TryGetSelectedListObject = Selection.ListObject
End Function

Public Function TryGetListObjectOnSheet() As ListObject
    If Activesheet.listobjects.Count = 1 Then
        Set TryGetListObjectOnSheet = Activesheet.listobjects(1)
    End If
End Function

Public Function TryGetOnlyListObjectInWorkbook() As ListObject
    Dim result As ListObject

    Dim ws As Worksheet
    For Each ws In Activeworkbook.Worksheets
        If ws.listobjects.Count = 1 Then
            ' Exit without returning a result if there is already a ListObject set
            If Not result Is Nothing Then Exit Function

            Set result = ws.listobjects(1)
        End If
    Next ws

    Set TryGetOnlyListObjectInWorkbook = result
End Function
