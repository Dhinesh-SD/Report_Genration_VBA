Attribute VB_Name = "nAVIGATION"


Sub scrollTonextPage()
With ActiveWindow
    If Cells(.ScrollRow, .ScrollColumn).Row < 4 Then
        ActiveWindow.ScrollRow = 59
    ElseIf Cells(.ScrollRow, .ScrollColumn).Row < 63 Then
        ActiveWindow.ScrollRow = 121
    End If
End With
End Sub

Sub scrollToPrevPage()
With ActiveWindow
    If Cells(.ScrollRow, .ScrollColumn).Row >= 73 Then
        ActiveWindow.ScrollRow = 59
    ElseIf Cells(.ScrollRow, .ScrollColumn).Row >= 12 Then
        ActiveWindow.ScrollRow = 1
    End If
End With
End Sub

