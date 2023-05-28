VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestAvailableColumns 
   Caption         =   "UserForm1"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4530
   OleObjectBlob   =   "frmTestAvailableColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestAvailableColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'@Folder "Test"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mViewModel As SplitTableViewModel
Attribute mViewModel.VB_VarHelpID = -1
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub chkShowUnsuitableColumns_Click()
    mViewModel.ShowUnsuitableColumns = Me.chkShowUnsuitableColumns.Value
End Sub

Private Sub chkShowHiddenColumns_Click()
    mViewModel.ShowHiddenColumns = Me.chkShowHiddenColumns.Value
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Sub lvAvailableColumns_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Not mViewModel.TrySelectColumnByName(Item.Text) Then
        Item.Checked = False
    End If
End Sub

Private Sub lvAvailableColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TrySelectColumnByName Item.Text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    InitalizeFromViewModel
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitalizeFromViewModel()
    mViewModel.AvailableColumns.InitializeListView Me.lvAvailableColumns
    mViewModel_PropertyChanged "SelectedListColumn"
    'mViewModel_PropertyChanged "UpdateListColumns"
End Sub

Private Sub mViewModel_PropertyChanged(ByVal PropertyName As String)
    Select Case PropertyName
        Case "SelectedListColumn":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "ShowUnsuitableColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "ShowHiddenColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
    End Select
    
    UpdateControls
End Sub

Private Sub UpdateControls()
    Me.chkShowUnsuitableColumns.Value = mViewModel.ShowUnsuitableColumns
    Me.chkShowHiddenColumns.Value = mViewModel.ShowHiddenColumns
End Sub
