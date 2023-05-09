Attribute VB_Name = "modGetListObject"
Option Explicit

Public Function GetListObject() As ListObject
    Set GetListObject = TryGetSelectedListObject()
    If Not GetListObject Is Nothing Then Exit Function

    Set GetListObject = TryGetListObjectOnSheet()
    If Not GetListObject Is Nothing Then Exit Function

    Set GetListObject = TryGetOnlyListObjectInWorkbook()
    If Not GetListObject Is Nothing Then Exit Function
End Function

'@Description "If there is a ListObject on the current Selection, return it. Otherwise, return Nothing."
Private Function TryGetSelectedListObject() As ListObject
    Set TryGetSelectedListObject = Selection.ListObject
End Function

'@Description "If there is only one ListObject on the current ActiveSheet, return it. Otherwise, return Nothing."
Private Function TryGetListObjectOnSheet() As ListObject
    If Activesheet.listobjects.Count = 1 Then
        Set TryGetListObjectOnSheet = Activesheet.listobjects(1)
    End If
End Function

'@Description "If there is only one ListObject in the entire ActiveWorkbook, return it. Otherwise, return Nothing."
Private Function TryGetOnlyListObjectInWorkbook() As ListObject
    Dim Result As ListObject

    Dim ws As Worksheet
    For Each ws In Activeworkbook.Worksheets
        If ws.listobjects.Count = 1 Then
            ' Exit without returning a result if there is already a ListObject set
            If Not Result Is Nothing Then Exit Function

            Set Result = ws.listobjects(1)
        End If
    Next ws

    Set TryGetOnlyListObjectInWorkbook = Result
End Function
