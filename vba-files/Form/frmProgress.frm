VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Splitting table..."
   ClientHeight    =   1515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4230
   OleObjectBlob   =   "frmProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Prompts"
Option Explicit

Private Const IN_PROGRESS_MSG As String = "Busy splitting your table into separate sheets..."
Private Const COMPLETE_MSG As String = "Busy splitting your table into separate sheets... done."

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Public Sub UpdateProgress(ByVal Percentage As Double)
    If (Percentage >= 1#) Then OnComplete
    Me.ProgressBar1.Value = Percentage
End Sub

Private Sub UserForm_Initialize()
    Me.cmdOK.Enabled = False
    Me.ProgressBar1.Max = 1#
    Me.ProgressBar1.Min = 0
    Me.Label1.Caption = IN_PROGRESS_MSG
End Sub

Private Sub OnComplete()
    Me.Label1.Caption = COMPLETE_MSG
    Me.cmdOK.Enabled = True
End Sub
