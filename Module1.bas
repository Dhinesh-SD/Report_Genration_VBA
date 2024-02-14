Attribute VB_Name = "Module1"
Sub RunSavedReport()
Dim ws As Worksheet
Dim table As ListObject
Set ws = Sheet6
Set table = ws.ListObjects(1)

Dim i As Long, col As Long
Dim ReportName As String

col = ws.Range("G1").Column

For i = 2 To ws.Cells(1, 1).End(xlDown).Row
    If ws.Cells(i, col).Value = True Then
        ReportName = ws.Cells(i, 2).Value
        For j = 2 To Sheet6.Range("A1").End(xlDown).Row
            UserForm2.ListBox1.AddItem Sheet6.Range("B" & i).Value
        Next j
        UserForm2.ListBox1.Selected(i - 2) = True
        UserForm2.CommandButton1.Value = True
    End If
Next i
End Sub
