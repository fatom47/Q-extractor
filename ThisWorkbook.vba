Option Explicit

' Occurs when the workbook is opened
Private Sub Workbook_Open()
    Call fceStopStart
End Sub

' Occurs before the workbook closes
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call fceStopStart
End Sub

Sub fceStopStart()
    Worksheets("Output").Cells.Clear
    Worksheets("Output").Cells.NumberFormat = "@"
    Worksheets("Input").Cells.Clear
    Worksheets("Input").Cells.NumberFormat = "@"
    Worksheets("Settings").Activate
    Worksheets("Settings").Range("A7:B8").Select
    'ActiveWorkbook.Save
End Sub
