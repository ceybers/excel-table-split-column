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

Event PropertyChanged(PropertyName As String, vNewValue As Object)

Private Type TState
    SelectedListObject As ListObject
    ListObjects As Collection
    SelectedListColumn As ListColumn
    ListColumns As Collection
    TargetSheets As Collection
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
    OnPropertyChanged "SelectedListColumn", This.SelectedListColumn
    
    UpdateTargetSheets
End Property

Public Property Get ListColumns() As Collection
    Debug.Assert Not This.ListColumns Is Nothing
    Set ListColumns = This.ListColumns
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
    OnPropertyChanged "SelectedListObject", This.SelectedListObject
    
    UpdateListColumns
End Property

Public Property Get ListObjects() As Collection
    If This.ListObjects Is Nothing Then Exit Property
    Set ListObjects = This.ListObjects
End Property

Public Sub Load(ByVal Workbook As Workbook)
    Set This.ListObjects = ListObjectHelpers.GetAllListObjects(Workbook)
    If This.ListObjects.Count > 0 Then
        Set SelectedListObject = This.ListObjects.Item(1)
    End If
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String, vNewValue As Object)
    Debug.Print "PropertyChanged('"; PropertyName; "', vNewValue)"
    RaiseEvent PropertyChanged(PropertyName, vNewValue)
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
    Dim ColumnAnalysis As ColumnAnalysis
    For Each ColumnAnalysis In This.ListColumns
        If ColumnAnalysis.ListColumn.Name = ListColumnName Then
            Set SelectedListColumn = ColumnAnalysis.ListColumn
            Exit Sub
        End If
    Next ColumnAnalysis
End Sub

Private Sub UpdateListColumns()
    Debug.Assert Not This.ListColumns Is Nothing
    Debug.Assert Not This.SelectedListObject Is Nothing
    
    CollectionHelpers.CollectionClear This.ListColumns
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.SelectedListObject.ListColumns
        Dim ColumnAnalysis As ColumnAnalysis
        Set ColumnAnalysis = New ColumnAnalysis
        ColumnAnalysis.Analyse ListColumn
        This.ListColumns.Add Key:=ListColumn.Name, Item:=ColumnAnalysis
    Next ListColumn
    
    OnPropertyChanged "UpdateListColumns", This.ListColumns 'TODO OnCollectionChanged?
    
    ' TODO Select first suitable column, not just first column
    SelectListColumnByName This.ListColumns.Item(4).ListColumn
End Sub

Private Sub Class_Initialize()
    Set This.ListColumns = New Collection
    Set This.TargetSheets = New Collection
End Sub

Private Sub UpdateTargetSheets()
    Dim ColumnAnalysis As ColumnAnalysis
    For Each ColumnAnalysis In This.ListColumns
        If ColumnAnalysis.ListColumn.Name = This.SelectedListColumn.Name Then
            Set This.TargetSheets = ColumnAnalysis.GetTargetSheets
            OnPropertyChanged "UpdateTargetSheets", This.TargetSheets 'TODO OnCollectionChanged?
            Exit Sub
        End If
    Next ColumnAnalysis
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
    OnPropertyChanged "UpdateTargetSheets", This.TargetSheets 'TODO OnCollectionChanged?
End Sub

Public Sub DoCheckNoTargetSheets()
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        TargetSheet.Used = False
    Next TargetSheet
    OnPropertyChanged "UpdateTargetSheets", This.TargetSheets 'TODO OnCollectionChanged?
End Sub

Public Sub TryCheckTargetSheet(ByVal TargetSheetName As String, ByVal vNewValue As Boolean)
    Dim TargetSheet As TargetSheet
    For Each TargetSheet In This.TargetSheets
        If TargetSheet.Name = TargetSheetName Then
            TargetSheet.Used = vNewValue
        End If
    Next TargetSheet
    OnPropertyChanged "UpdateTargetSheets", This.TargetSheets 'TODO OnCollectionChanged?
End Sub
