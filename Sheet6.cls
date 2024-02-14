VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim i As Long

For i = 2 To Sheet6.Range("A1").End(xlDown).Row
    UserForm2.ListBox1.AddItem Sheet6.Range("B" & i).Value
Next i
With UserForm2
    .Top = Application.Top + Sheet6.Range("A10").Top
    .Left = Application.Left + Sheet6.Range("C10").Left

End With
UserForm2.Show vbModeless
End Sub
