VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SplitTableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM"
Option Explicit

Event PropertyChanged(PropertyName As String)

Private Type TState
    SelectedListObject As ListObject
    ListObjects As Collection
    SelectedListColumn As ListColumn
    ListColumns As Collection ' TODO Remove ifo Available Columns
    AvailableColumns As AvailableColumns
    TargetSheets As Collection ' of type TargetSheet
    ShowHiddenColumns As Boolean
    ShowUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    If SelectedListObject Is Nothing Then Exit Property
    If SelectedListColumn Is Nothing Then Exit Property
    If TargetSheetPercentage = 0 Then Exit Property
    
    IsValid = True
End Property

Public Property Get CanSelectAll() As Boolean
    CanSelectAll = TargetSheetPercentage < 1 And This.TargetSheets.Count > 0
End Property

Public Property Get CanSelectNone() As Boolean
    CanSelectNone = TargetSheetPercentage > 0
End Property

Public Property Get TargetSheets() As Collection
    Set TargetSheets = This.TargetSheets
End Property

Public Property Get SelectedListColumn() As ListColumn
    Set SelectedListColumn = This.SelectedListColumn
End Property

Public Property Set SelectedListColumn(ByVal vNewValue As ListColumn)
    If vNewValue Is Nothing Then Exit Property
    
    If Not This.SelectedListColumn Is Nothing Then
        If This.SelectedListColumn = vNewValue Then
            Exit Property
        End If
    End If
    
    Set This.SelectedListColumn = vNewValue
    OnPropertyChanged "SelectedListColumn"
    
    UpdateTargetSheets
End Property

Public Property Get AvailableColumns() As AvailableColumns
    Set AvailableColumns = This.AvailableColumns
End Property

Public Property Get SelectedListObject() As ListObject
    If This.SelectedListObject Is Nothing Then Exit Property
    Set SelectedListObject = This.SelectedListObject
End Property
  
Public Property Set SelectedListObject(ByVal vNewValue As ListObject)
    If vNewValue Is Nothing Then Exit Property
    
    If Not This.SelectedListObject Is Nothing Then
        If This.SelectedListObject = vNewValue Then
            Exit Property
        End If
    End If
    
    Set This.SelectedListObject = vNewValue
    OnPropertyChanged "SelectedListObject"
    
    UpdateAvailableColumns
End Property

Public Property Get ListObjects() As Collection
    If This.ListObjects Is Nothing Then Exit Property
    Set ListObjects = This.ListObjects
End Property

Public Property Get ShowHiddenColumns() As Boolean
    ShowHiddenColumns = This.ShowHiddenColumns
End Property

Public Property Let ShowHiddenColumns(ByVal vNewValue As Boolean)
    If This.ShowHiddenColumns = vNewValue Then Exit Property
    This.ShowHiddenColumns = vNewValue
    This.AvailableColumns.ShowHidden = This.ShowHiddenColumns
    OnPropertyChanged "ShowHiddenColumns"
End Property

Public Property Get ShowUnsuitableColumns() As Boolean
    ShowUnsuitableColumns = This.ShowUnsuitableColumns
End Property

Public Property Let ShowUnsuitableColumns(ByVal vNewValue As Boolean)
    If This.ShowUnsuitableColumns = vNewValue Then Exit Property
    This.ShowUnsuitableColumns = vNewValue
    This.AvailableColumns.ShowUnsuitable = This.ShowUnsuitableColumns
    OnPropertyChanged "ShowUnsuitableColumns"
End Property

Public Property Get DeleteExistingSheets() As Boolean
    DeleteExistingSheets = This.DeleteExistingSheets
End Property

Public Property Let DeleteExistingSheets(ByVal vNewValue As Boolean)
    If This.DeleteExistingSheets = vNewValue Then Exit Property
    This.DeleteExistingSheets = vNewValue
    OnPropertyChanged "DeleteExistingSheets"
End Property

Public Property Get RemoveOtherSheets() As Boolean
    RemoveOtherSheets = This.RemoveOtherSheets
End Property

Public Property Let RemoveOtherSheets(ByVal vNewValue As Boolean)
    If This.RemoveOtherSheets = vNewValue Then Exit Property
    This.RemoveOtherSheets = vNewValue
    OnPropertyChanged "RemoveOtherSheets"
End Property

Public Sub Load(ByVal Workbook As Workbook)
    Set This.ListObjects = ListObjectHelpers.GetAllListObjects(Workbook)
    If This.ListObjects.Count > 0 Then
        Set SelectedListObject = This.ListObjects.Item(1)
    End If
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    Debug.Print "PropertyChanged('"; PropertyName; "')"
    RaiseEvent PropertyChanged(PropertyName)
End Sub

Public Sub SelectListObjectByName(ByVal ListObjectName As String)
    Dim ListObject As ListObject
    For Each ListObject In This.ListObjects
        If ListObject.Name = ListObjectName Then
            Set SelectedListObject = ListObject
            Exit Sub
        End If
    Next ListObject
End Sub

Public Sub SelectListColumnByName(ByVal ListColumnName As String)
    Set SelectedListColumn = This.SelectedListObject.ListColumns(ListColumnName)
End Sub

Private Sub UpdateAvailableColumns()
    Debug.Assert Not This.AvailableColumns Is Nothing
    Debug.Assert Not This.SelectedListObject Is Nothing
    
    This.AvailableColumns.Load This.SelectedListObject
    
    OnPropertyChanged "UpdateListColumns"

    ' TODO Select first suitable column, not just first column
    SelectListColumnByName This.AvailableColumns.GetAvailableColumns()(3).Name
    
    'CollectionHelpers.CollectionClear This.ListColumns
    
    'Dim ListColumn As ListColumn
    'For Each ListColumn In This.SelectedListObject.ListColumns
    '    Dim ColumnAnalysis As ColumnAnalysis
    '    Set ColumnAnalysis = New ColumnAnalysis
    '    ColumnAnalysis.Analyse ListColumn
    '    This.ListColumns.Add Key:=ListColumn.Name, Item:=ColumnAnalysis
    'Next ListColumn
    '
    'OnPropertyChanged "UpdateListColumns" 'TODO OnCollectionChanged?
    
End Sub

Private Sub Class_Initialize()
    'Set This.ListColumns = New Collection
    Set This.AvailableColumns = New AvailableColumns
    Set This.TargetSheets = New Collection
End Sub

Private Sub UpdateTargetSheets()
    Set This.TargetSheets = ColumnAnalysis.Create(This.SelectedListColumn).GetTargetSheets
    OnPropertyChanged "UpdateTargetSheets"
End Sub

Private Function TargetSheetPercentage() As Double
    Dim Numerator As Long
    Dim Denominator As Long
    Denominator = This.TargetSheets.Count
    If Denominator = 0 Then Exit Function
    
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        If TargetSheet.Used Then
            Numerator = Numerator + 1
        End If
    Next TargetSheet
    
    TargetSheetPercentage = Numerator / Denominator
End Function

Public Sub DoCheckAllTargetSheets()
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        TargetSheet.Used = True
    Next TargetSheet
    OnPropertyChanged "UpdateTargetSheets" 'TODO OnCollectionChanged?
End Sub

Public Sub DoCheckNoTargetSheets()
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        TargetSheet.Used = False
    Next TargetSheet
    OnPropertyChanged "UpdateTargetSheets" 'TODO OnCollectionChanged?
End Sub

Public Sub TryCheckTargetSheet(ByVal TargetSheetName As String, ByVal vNewValue As Boolean)
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        If TargetSheet.Name = TargetSheetName Then
            TargetSheet.Used = vNewValue
        End If
    Next TargetSheet
    OnPropertyChanged "UpdateTargetSheets" 'TODO OnCollectionChanged?
End Sub
