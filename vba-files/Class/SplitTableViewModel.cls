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
    AvailableColumns As AvailableColumns
    TargetSheets As TargetSheets
    ShowHiddenColumns As Boolean
    ShowUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    If SelectedListObject Is Nothing Then Exit Property
    If This.AvailableColumns.Selected Is Nothing Then Exit Property
    If CanSelectNone = False Then Exit Property
    
    IsValid = True
End Property

Public Property Get CanSelectAll() As Boolean
    CanSelectAll = This.TargetSheets.CanSelectAll
End Property

Public Property Get CanSelectNone() As Boolean
    CanSelectNone = This.TargetSheets.CanSelectNone
End Property

Public Property Get TargetSheets() As TargetSheets
    Set TargetSheets = This.TargetSheets
End Property

Public Property Get AvailableColumns() As AvailableColumns
    Set AvailableColumns = This.AvailableColumns
End Property

Public Property Get SelectedListObject() As ListObject
    If This.SelectedListObject Is Nothing Then Exit Property
    Set SelectedListObject = This.SelectedListObject
End Property
  
Public Property Set SelectedListObject(ByVal vNewvalue As ListObject)
    If vNewvalue Is Nothing Then Exit Property
    
    If Not This.SelectedListObject Is Nothing Then
        If This.SelectedListObject = vNewvalue Then
            Exit Property
        End If
    End If
    
    Set This.SelectedListObject = vNewvalue
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

Public Property Let ShowHiddenColumns(ByVal vNewvalue As Boolean)
    If This.ShowHiddenColumns = vNewvalue Then Exit Property
    This.ShowHiddenColumns = vNewvalue
    This.AvailableColumns.ShowHidden = This.ShowHiddenColumns
    OnPropertyChanged "ShowHiddenColumns"
End Property

Public Property Get ShowUnsuitableColumns() As Boolean
    ShowUnsuitableColumns = This.ShowUnsuitableColumns
End Property

Public Property Let ShowUnsuitableColumns(ByVal vNewvalue As Boolean)
    If This.ShowUnsuitableColumns = vNewvalue Then Exit Property
    This.ShowUnsuitableColumns = vNewvalue
    This.AvailableColumns.ShowUnsuitable = This.ShowUnsuitableColumns
    OnPropertyChanged "ShowUnsuitableColumns"
End Property

Public Property Get DeleteExistingSheets() As Boolean
    DeleteExistingSheets = This.DeleteExistingSheets
End Property

Public Property Let DeleteExistingSheets(ByVal vNewvalue As Boolean)
    If This.DeleteExistingSheets = vNewvalue Then Exit Property
    This.DeleteExistingSheets = vNewvalue
    OnPropertyChanged "DeleteExistingSheets"
End Property

Public Property Get RemoveOtherSheets() As Boolean
    RemoveOtherSheets = This.RemoveOtherSheets
End Property

Public Property Let RemoveOtherSheets(ByVal vNewvalue As Boolean)
    If This.RemoveOtherSheets = vNewvalue Then Exit Property
    This.RemoveOtherSheets = vNewvalue
    OnPropertyChanged "RemoveOtherSheets"
End Property

Private Sub Class_Initialize()
    Set This.AvailableColumns = New AvailableColumns
    Set This.TargetSheets = New TargetSheets
End Sub

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
    'Set SelectedListColumn = This.SelectedListObject.ListColumns(ListColumnName)
    If This.AvailableColumns.TrySelectByName(ListColumnName) Then
        OnPropertyChanged "SelectedListColumn"
    End If
End Sub

Private Sub UpdateAvailableColumns()
    Debug.Assert Not This.AvailableColumns Is Nothing
    Debug.Assert Not This.SelectedListObject Is Nothing
    
    This.AvailableColumns.Load This.SelectedListObject
    
    OnPropertyChanged "UpdateListColumns"

    SelectListColumnByName This.AvailableColumns.GetAvailableColumns()(1).Name
End Sub

Private Sub UpdateTargetSheets()
    'Debug.Assert Not This.SelectedListColumn Is Nothing
    This.TargetSheets.Load This.AvailableColumns.Selected
    'Set This.TargetSheets = ColumnAnalysis.Create(This.SelectedListColumn).GetTargetSheets
    OnPropertyChanged "UpdateTargetSheets"
End Sub

Public Sub TryCheckTargetSheet(ByVal TargetSheetName As String, ByVal vNewvalue As Boolean)
    If This.TargetSheets.TryCheckByName(TargetSheetName, vNewvalue) Then
        OnPropertyChanged "UpdateTargetSheets" 'TODO OnCollectionChanged?
    End If
End Sub

Public Function TrySelectColumnByName(ByVal ColumnName As String) As Boolean
    If This.AvailableColumns.TrySelectByName(ColumnName) Then
        TrySelectColumnByName = True
        OnPropertyChanged "SelectedListColumn"
    End If
End Function
