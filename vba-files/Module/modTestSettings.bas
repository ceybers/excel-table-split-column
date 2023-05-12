Attribute VB_Name = "modTestSettings"
'@Folder "Test"
Option Explicit

Public Sub TestSettings()
    Dim Settings As ISettings
    Set Settings = New MyDocsSettings
    Settings.Load
    Debug.Print "Get SHOWHIDDEN = "; Settings.GetFlag("SHOWHIDDEN")
    Settings.SetFlag "SHOWHIDDEN", "FALSE"
End Sub
