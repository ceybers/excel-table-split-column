VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorkbookProtected 
   Caption         =   "Workbook is Protected"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "frmWorkbookProtected.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWorkbookProtected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IMAGEMSO_SIZE As Long = 32

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Set Me.Image1.Picture = Application.CommandBars.GetImageMso("ReviewProtectWorkbook", IMAGEMSO_SIZE, IMAGEMSO_SIZE)
End Sub

