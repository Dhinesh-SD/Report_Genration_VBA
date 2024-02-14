Attribute VB_Name = "worksheet_protection"
Option Explicit
 Dim ws As Worksheet, a As Range
Sub protectWorksheet()

For Each ws In ThisWorkbook.Worksheets
    protc ws
Next ws
End Sub

Sub unProtectWorksheet()
'
For Each ws In ThisWorkbook.Worksheets
    unprotc ws
Next ws

End Sub
Sub unProtectWorkbook(Optional booksA As Workbook)
'
Dim count As Long, maxsheets As Long
count = 0
maxsheets = booksA.Sheets.count
For Each ws In booksA.Worksheets
    unprotc ws
    count = count + 1
Next ws
End Sub

Sub ProtectWorkbook(Optional booksA As Workbook)
'
Dim count As Long, maxsheets As Long
count = 0
maxsheets = booksA.Sheets.count
For Each ws In booksA.Worksheets
    protc ws
    count = count + 1
Next ws
End Sub
Sub onr()
unprotc
End Sub

Sub unprotc(Optional Active As Worksheet)
Application.ScreenUpdating = False
If Active Is Nothing Then
    Set Active = ActiveSheet
End If
Active.Unprotect Password:=""
Application.ScreenUpdating = True
End Sub

Sub protc(Optional Active As Worksheet)
Application.ScreenUpdating = False
If Active Is Nothing Then
    Set Active = ActiveSheet
End If
Active.Protect Password:="", DrawingObjects:=True, UserInterfaceonly:=True, Contents:=True, Scenarios:= _
        False, AllowSorting:=True, AllowFiltering:=True
Application.ScreenUpdating = True
End Sub





