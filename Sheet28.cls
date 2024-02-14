VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Shapes("VST").Visible = msoFalse
End Sub


Private Sub Worksheet_Activate()
ActiveWindow.Zoom = 100

If Left(Right(Range("A1").Value, 2), 1) <> "C" Then
    ActiveWindow.ScrollRow = Range("A120").Row
    Debug.Print Left(Right(Range("A1").Value, 2), 1)
    Exit Sub
End If

If Right(Range("A1").Value, 1) > 5 Then
    ActiveWindow.ScrollRow = Range("A59").Row
ElseIf Right(Range("A1").Value, 1) < 5 Then
    ActiveWindow.ScrollRow = Range("A1").Row
End If
End Sub
