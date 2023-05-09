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
    ListColumns As Collection
End Type
Private This As TState

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
End Sub

Private Sub Class_Initialize()
    Set This.ListColumns = New Collection
End Sub
