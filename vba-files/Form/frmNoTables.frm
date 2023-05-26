VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNoTables 
   Caption         =   "No Tables Found"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "frmNoTables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNoTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MVVM.TableSplit.Views"
Option Explicit

Private Const IMAGEMSO_SIZE As Long = 32

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Set Me.Image1.Picture = Application.CommandBars.GetImageMso("TableInsertDialog", IMAGEMSO_SIZE, IMAGEMSO_SIZE)
End Sub
