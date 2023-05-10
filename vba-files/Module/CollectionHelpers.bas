Attribute VB_Name = "CollectionHelpers"
'@Folder "Helpers"
Option Explicit

'@Description "Returns True if the given Value exists in a Collection."
Public Function ExistsInCollection(ByVal Collection As Object, ByVal Value As Variant) As Boolean
Attribute ExistsInCollection.VB_Description = "Returns True if the given Value exists in a Collection."
    Debug.Assert Not Collection Is Nothing
    
    Dim ThisValue As Variant
    For Each ThisValue In Collection
        If ThisValue = Value Then
            ExistsInCollection = True
            Exit Function
        End If
    Next ThisValue
End Function

'@Description "Removes all items in a Collection."
Public Sub CollectionClear(ByVal Collection As Collection)
Attribute CollectionClear.VB_Description = "Removes all items in a Collection."
    Debug.Assert Not Collection Is Nothing
    
    Dim i As Long
    For i = Collection.Count To 1 Step -1
        Collection.Remove i
    Next i
End Sub
