Attribute VB_Name = "modProcessSplitTableVM"
'@Folder("Test")
Option Explicit


Public Sub ProcessViewModel(ByVal ViewModel As SplitTableViewModel)
    If ViewModel.RemoveOtherSheets Then
        DoRemoveOtherSheets ViewModel.AvailableTables.Selected
    End If
    
    If ViewModel.DeleteExistingSheets Then
        DoRemoveExistingSheets ViewModel.AvailableTables.Selected, ViewModel.TargetSheets.GetSelectedSheetNames
    End If
    
    SplitTable ViewModel.AvailableColumns.Selected, ViewModel.TargetSheets.GetSelectedSheetNames
End Sub

Public Sub DoRemoveOtherSheets(ByVal ListObject As ListObject)
    Dim KeepWorksheet As Worksheet
    Set KeepWorksheet = ListObject.Parent
    
    Dim Worksheet As Worksheet
    For Each Worksheet In KeepWorksheet.Parent.Worksheets
        If Worksheet.Name <> KeepWorksheet.Name Then
            Application.DisplayAlerts = False
            Worksheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Worksheet
End Sub

Public Sub DoRemoveExistingSheets(ByVal ListObject As ListObject, ByVal SheetNames As Collection)
    Dim Worksheet As Worksheet
    For Each Worksheet In ListObject.Parent.Parent.Worksheets
        If ExistsInCollection(SheetNames, Worksheet.Name) Then
            Application.DisplayAlerts = False
            Worksheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Worksheet
End Sub
