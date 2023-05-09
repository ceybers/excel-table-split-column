Attribute VB_Name = "CollectionHelpers"
'@Folder "Helpers"
Option Explicit

Public Function ExistsInCollection(ByVal coll As Collection, ByVal val As Variant) As Boolean
    Debug.Assert Not coll Is Nothing
    
    Dim v As Variant
    
    For Each v In coll
        If v = val Then
            ExistsInCollection = True
            Exit Function
        End If
    Next v
End Function

Public Sub CollectionClear(ByVal coll As Collection)
    Dim i As Long
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
End Sub


