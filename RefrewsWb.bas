Attribute VB_Name = "RefrewsWb"
Option Explicit
Sub RefreshTables()

Dim Obj As Object, folder As Object, file As Variant
Dim folderLocation As String
Dim i As Integer, lastrow As Integer

lastrow = Sheet32.Cells(1, 1).End(xlDown).Row
folderLocation = "I:\Office of Procurement and Grants\Team - Services\Reports\"

'Set folder = Obj.getfolder(folderLocation)
Dim diff As Date
Dim StrFile As String
    'Debug.Print "in LoopThroughFiles. inputDirectoryToScanForFile: ", inputDirectoryToScanForFile
Set Obj = CreateObject("Scripting.FileSystemObject")
StrFile = Dir(folderLocation & "*.xlsx")
    Do While Len(StrFile) > 0
        For i = 2 To lastrow
            If Replace(StrFile, Right(StrFile, 5), "") = Replace(Sheet32.Cells(i, 1).Value, Right(Sheet32.Cells(i, 1).Value, 5), "") Then
                diff = Obj.GetFile(folderLocation & StrFile).DateCreated - Sheet32.Cells(i, 5).Value
                If Hour(diff) > 0 Then
                    GoTo SkipHere
                End If
            End If
        Next i
        StrFile = Dir
    Loop
MsgBox "Reports are upto date! Download New reports To update Tables"
Exit Sub
SkipHere:
ThisWorkbook.RefreshAll
MsgBox "Refreshed Reports!"
Sheet32.Range("UpdateTime").Value = Now
End Sub

Sub rfrsh()
    ThisWorkbook.RefreshAll
End Sub
