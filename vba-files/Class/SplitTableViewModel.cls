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

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    'SelectedListObject As ListObject
    'ListObjects As Collection
    AvailableTables As AvailableTables
    AvailableColumns As AvailableColumns
    TargetSheets As TargetSheets
    ShowHiddenColumns As Boolean
    ShowUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
    Settings As MyDocsSettings
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    If This.AvailableTables.Selected Is Nothing Then Exit Property
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

Public Property Get AvailableTables() As AvailableTables
    Set AvailableTables = This.AvailableTables
End Property

Public Property Get TargetSheets() As TargetSheets
    Set TargetSheets = This.TargetSheets
End Property

Public Property Get AvailableColumns() As AvailableColumns
    Set AvailableColumns = This.AvailableColumns
End Property

Public Property Get ShowHiddenColumns() As Boolean
    ShowHiddenColumns = This.ShowHiddenColumns
End Property

Public Property Let ShowHiddenColumns(ByVal vNewValue As Boolean)
    If This.ShowHiddenColumns = vNewValue Then Exit Property
    This.ShowHiddenColumns = vNewValue
    This.AvailableColumns.ShowHidden = This.ShowHiddenColumns
    OnPropertyChanged "ShowHiddenColumns"
    SaveSettings
End Property

Public Property Get ShowUnsuitableColumns() As Boolean
    ShowUnsuitableColumns = This.ShowUnsuitableColumns
End Property

Public Property Let ShowUnsuitableColumns(ByVal vNewValue As Boolean)
    If This.ShowUnsuitableColumns = vNewValue Then Exit Property
    This.ShowUnsuitableColumns = vNewValue
    This.AvailableColumns.ShowUnsuitable = This.ShowUnsuitableColumns
    OnPropertyChanged "ShowUnsuitableColumns"
    SaveSettings
End Property

Public Property Get DeleteExistingSheets() As Boolean
    DeleteExistingSheets = This.DeleteExistingSheets
End Property

Public Property Let DeleteExistingSheets(ByVal vNewValue As Boolean)
    If This.DeleteExistingSheets = vNewValue Then Exit Property
    This.DeleteExistingSheets = vNewValue
    OnPropertyChanged "DeleteExistingSheets"
    SaveSettings
End Property

Public Property Get RemoveOtherSheets() As Boolean
    RemoveOtherSheets = This.RemoveOtherSheets
End Property

Public Property Let RemoveOtherSheets(ByVal vNewValue As Boolean)
    If This.RemoveOtherSheets = vNewValue Then Exit Property
    This.RemoveOtherSheets = vNewValue
    OnPropertyChanged "RemoveOtherSheets"
    SaveSettings
End Property

Private Sub Class_Initialize()
    Set This.AvailableTables = New AvailableTables
    Set This.AvailableColumns = New AvailableColumns
    Set This.TargetSheets = New TargetSheets
    Set This.Settings = New MyDocsSettings
    This.Settings.Load
End Sub

Public Sub Load(ByVal Workbook As Workbook)
    This.AvailableTables.Load Workbook
    If Not This.AvailableTables.Selected Is Nothing Then
        This.AvailableColumns.Load This.AvailableTables.Selected
    End If
    LoadSettings
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    'Debug.Print "PropertyChanged('"; PropertyName; "')"
    RaiseEvent PropertyChanged(PropertyName)
End Sub

Private Sub UpdateAvailableColumns()
    Debug.Assert Not This.AvailableColumns Is Nothing
    Debug.Assert Not This.AvailableTables.Selected Is Nothing
    
    This.AvailableColumns.Load This.AvailableTables.Selected
    
    OnPropertyChanged "UpdateListColumns"

    'TrySelectColumnByName This.AvailableColumns.GetAvailableColumns()(1).Name
End Sub

Private Sub UpdateTargetSheets()
    This.TargetSheets.Load This.AvailableColumns.Selected
    OnPropertyChanged "UpdateTargetSheets"
End Sub

Public Sub TryCheckTargetSheet(ByVal TargetSheetName As String, ByVal vNewValue As Boolean)
    If This.TargetSheets.TryCheckByName(TargetSheetName, vNewValue) Then
        OnPropertyChanged "UpdateTargetSheets"
    End If
End Sub

Public Function TrySelectColumnByName(ByVal ColumnName As String) As Boolean
    If This.AvailableColumns.TrySelectByName(ColumnName) Then
        TrySelectColumnByName = True
        OnPropertyChanged "SelectedListColumn"
        UpdateTargetSheets
    End If
End Function

Public Function TrySelectTableByName(ByVal TableName As String) As Boolean
    If This.AvailableTables.TrySelectByName(TableName) Then
        TrySelectTableByName = True
        OnPropertyChanged "SelectedListObject"
        UpdateAvailableColumns
    End If
End Function

Private Sub LoadSettings()
    This.ShowHiddenColumns = This.Settings.GetFlag(SHOW_HIDDEN_FLAG)
    This.ShowUnsuitableColumns = This.Settings.GetFlag(SHOW_UNSUITABLE_FLAG)
    This.DeleteExistingSheets = This.Settings.GetFlag(REPLACE_EXISTING_FLAG)
    This.RemoveOtherSheets = This.Settings.GetFlag(REPLACE_OTHER_FLAG)
End Sub

Private Sub SaveSettings()
    This.Settings.SetFlag SHOW_HIDDEN_FLAG, This.ShowHiddenColumns
    This.Settings.SetFlag SHOW_UNSUITABLE_FLAG, This.ShowUnsuitableColumns
    This.Settings.SetFlag REPLACE_EXISTING_FLAG, This.DeleteExistingSheets
    This.Settings.SetFlag REPLACE_OTHER_FLAG, This.RemoveOtherSheets
    This.Settings.Save
End Sub
