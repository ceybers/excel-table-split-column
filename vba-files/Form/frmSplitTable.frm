VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitTable 
   Caption         =   "Split Table by Columns"
   ClientHeight    =   8445.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250.001
   OleObjectBlob   =   "frmSplitTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplitTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MVVM.TableSplit.Views"
Option Explicit
Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mViewModel As SplitTableViewModel
Attribute mViewModel.VB_VarHelpID = -1
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cboTable_Change()
    mViewModel.TrySelectTableByName Me.cboTable.Text
End Sub

Private Sub chkDeleteExistingSheets_Click()
    mViewModel.DeleteExistingSheets = Me.chkDeleteExistingSheets.Value
End Sub

Private Sub chkShowUnsuitableColumns_Click()
    mViewModel.ShowUnsuitableColumns = Me.chkShowUnsuitableColumns.Value
End Sub

Private Sub chkRemoveOtherSheets_Click()
    mViewModel.RemoveOtherSheets = Me.chkRemoveOtherSheets.Value
    Me.chkDeleteExistingSheets.Enabled = Not mViewModel.RemoveOtherSheets
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

Private Sub cmbSelectAll_Click()
    If mViewModel.TargetSheets.SelectAll Then
        mViewModel_PropertyChanged "UpdateTargetSheets"
    End If
End Sub

Private Sub cmbSelectNone_Click()
    If mViewModel.TargetSheets.SelectNone Then
        mViewModel_PropertyChanged "UpdateTargetSheets"
    End If
End Sub

Private Sub lvAvailableColumns_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Not mViewModel.TrySelectColumnByName(Item.Text) Then
        Item.Checked = False
    End If
End Sub

Private Sub lvAvailableColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TrySelectColumnByName Item.Text
End Sub

Private Sub lvUsedValues_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TryCheckTargetSheet Item.Text, Item.Checked
End Sub

Private Sub lvUsedValues_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mViewModel.TryCheckTargetSheet Item.Text, Item.Checked
End Sub

Private Sub SetLabelPictures()
    Set Me.lblPicColumns.Picture = Application.CommandBars.GetImageMso("TableColumnSelect", 32, 32)
    Set Me.lblPicSheets.Picture = Application.CommandBars.GetImageMso("HeaderFooterSheetNameInsert", 32, 32)
    Set Me.lblPicTables.Picture = Application.CommandBars.GetImageMso("TableAutoFormat", 32, 32)
    Set Me.lblPicSettings.Picture = Application.CommandBars.GetImageMso("OmsViewAccountSetting", 32, 32)
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
    
    SetLabelPictures
    InitalizeFromViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub mViewModel_PropertyChanged(ByVal PropertyName As String)
    Select Case PropertyName
        Case "SelectedListObject":
            mViewModel.AvailableTables.UpdateCombobox Me.cboTable
            mViewModel.AvailableTables.Activate
        Case "SelectedListColumn":
            ' Updates the LHS ListView with the correct checkboxes
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
            mViewModel.TargetSheets.InitalizeTargetSheets Me.lvUsedValues
        Case "ResetAvailableColumns":
            ' Updates the LHS ListView with correct items when switch tables
            mViewModel.AvailableColumns.InitializeListView Me.lvAvailableColumns
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "UpdateTargetSheets":
            mViewModel.AvailableColumns.Activate
            mViewModel.TargetSheets.InitalizeTargetSheets Me.lvUsedValues
            mViewModel.TargetSheets.UpdateListView Me.lvUsedValues
        Case "ShowUnsuitableColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
        Case "ShowHiddenColumns":
            mViewModel.AvailableColumns.UpdateListView Me.lvAvailableColumns
    End Select
    
    UpdateControls
End Sub

Private Sub InitalizeFromViewModel()
    UpdateControls

    mViewModel.AvailableTables.InitializeCombobox Me.cboTable
    mViewModel.AvailableColumns.ShowHidden = mViewModel.ShowHiddenColumns
    mViewModel.AvailableColumns.ShowUnsuitable = mViewModel.ShowUnsuitableColumns
    mViewModel.AvailableColumns.InitializeListView Me.lvAvailableColumns
    mViewModel.TargetSheets.InitalizeTargetSheets Me.lvUsedValues
    
    mViewModel_PropertyChanged "SelectedListObject"
    mViewModel_PropertyChanged "SelectedListColumn"
    mViewModel_PropertyChanged "UpdateTargetSheets"
End Sub

Private Sub UpdateControls()
    Me.cmbSelectAll.Enabled = mViewModel.CanSelectAll
    Me.cmbSelectNone.Enabled = mViewModel.CanSelectNone
    Me.cmbOK.Enabled = mViewModel.IsValid
    
    Me.chkDeleteExistingSheets.Value = mViewModel.DeleteExistingSheets
    Me.chkShowUnsuitableColumns.Value = mViewModel.ShowUnsuitableColumns
    Me.chkRemoveOtherSheets.Value = mViewModel.RemoveOtherSheets
    Me.chkShowHiddenColumns.Value = mViewModel.ShowHiddenColumns
End Sub
