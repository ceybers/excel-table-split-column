Attribute VB_Name = "modTestProgressBar"
Option Explicit

Public Sub TestProgressBar()
    With frmProgress
        .Show
        WaitOneSecond
        .UpdateProgress 0.25
        WaitOneSecond
        .UpdateProgress 0.75
        WaitOneSecond
        .UpdateProgress 1#
        
    End With
End Sub

Private Sub WaitOneSecond()
    Application.Wait (Now + TimeValue("00:00:01"))
End Sub
