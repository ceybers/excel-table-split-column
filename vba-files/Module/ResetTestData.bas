Attribute VB_Name = "modResetTestData"
Option Explicit

Private Const TEST_DATA_SHEETNAME as string = "Top10"
Private Const TEST_DATA_NAME as string = "TestData.xlsx"
Private Const TEST_DATA_FULLNAME as string = "C:\Users\User\Documents\Work\excel-table-split-column\TestData.xlsx"

Public Sub ResetTestData()
    TryOpenTestDataWorkbook
    RemoveExcessWorksheets
    CopyTestDataWorksheet
    RemoveLastWorksheet
    CloseTestDataWorkbook
End Sub

Private Sub TryOpenTestDataWorkbook()
    Dim wb as Workbook
    For each wb in Workbooks
        If wb.name = TEST_DATA_NAME then
            Exit Sub
        End if
        Next

        Workbooks.Open Filename:=TEST_DATA_FULLNAME, ReadOnly:=True
End Sub

Private Sub RemoveExcessWorksheets()
    Dim i As Long
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count to 2 Step -1
        ThisWorkbook.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets(1).Name = "not" & TEST_DATA_SHEETNAME
End Sub

Private Sub CopyTestDataWorksheet()
    Workbooks(TEST_DATA_NAME).Worksheets(TEST_DATA_SHEETNAME).Copy ThisWorkbook.Worksheets(1)
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Worksheets(1).Range("A2").Select
End Sub

Private Sub RemoveLastWorksheet()
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(2).Delete
    Application.DisplayAlerts = True
End Sub

Private Sub CloseTestDataWorkbook()
    Workbooks(TEST_DATA_NAME).Close SaveChanges:=False
End Sub