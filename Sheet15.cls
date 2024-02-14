VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
ActiveWindow.Zoom = 100
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Shapes("VST").Visible = msoFalse
End Sub

