VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4500
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6096
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
Dim i As Long, j As Long
Dim DsSheet As Worksheet
Dim Fields() As String
Dim Filter_1 As Long
Dim Filter_2 As Long
Dim Filter_3 As Long
Dim Filter_4 As Long
Dim cntrls As Variant
If Me.ListBox1.List(0) = "" Then Exit Sub

For i = 0 To Me.ListBox1.ListCount - 1
     If Me.ListBox1.Selected(i) = True Then
        Set DsSheet = ThisWorkbook.Worksheets(Sheet6.Range("C" & i + 2).Value)
        DsSheet.Range("A1").Value = Trim(Sheet6.Range("D" & i + 2).Value)
        DsSheet.Activate
        drillDownButton
        Fields = Split(Sheet6.Range("E" & i + 2).Value, ",")
        For j = LBound(Fields) To UBound(Fields)
            UserForm1.ListBox2.AddItem Fields(j)
        Next j
        UserForm1.ReportName.Value = Sheet6.Range("B" & i + 2).Value
        UserForm1.TextBox1.Value = Sheet6.Range("F" & i + 2).Value
        UserForm1.TextBox3.Value = Sheet6.Range("H" & i + 2).Value
        UserForm1.ChartPage.Caption = Sheet6.Range("C" & i + 2).Value
        UserForm1.ChartTable.Caption = Sheet6.Range("D" & i + 2).Value
        UserForm1.Label14.Caption = Sheet6.Range("J" & i + 2).Value
        'UserForm1.Label14.Caption
Dim filters() As String

        If Sheet6.Range("I" & i + 2).Value <> "" Then
            filters = Split(Sheet6.Range("I" & i + 2).Value, ";")
            For j = LBound(filters) To UBound(filters)
                UserForm1.ListBox3.AddItem filters(j)
            Next j
        End If
        UserForm1.CommandButton3.Value = True
     End If
Next i
Unload Me
Sheet6.Select
End Sub

Private Sub CommandButton2_Click()
Dim i As Long, j As Long
Dim DsSheet As Worksheet
Dim Fields() As String
Dim Filter_1 As Long
Dim Filter_2 As Long
Dim Filter_3 As Long
Dim Filter_4 As Long
Dim cntrls As Variant
Dim yesno As String
Dim k As Long
If Me.ListBox1.List(0) = "" Then Exit Sub

For i = 0 To Me.ListBox1.ListCount - 1
     If Me.ListBox1.Selected(i) = True Then
        yesno = MsgBox("Would you like to edit this report? :" & Me.ListBox1.List(i), vbYesNo, "Confirmation Dialouge Box")
        If yesno = vbYes Then
            Set DsSheet = ThisWorkbook.Worksheets(Sheet6.Range("C" & i + 2).Value)
            DsSheet.Range("A1").Value = Trim(Sheet6.Range("D" & i + 2).Value)
            DsSheet.Activate
            drillDownButton
            Fields = Split(Sheet6.Range("E" & i + 2).Value, ",")
            For j = LBound(Fields) To UBound(Fields)
                UserForm1.ListBox2.AddItem Fields(j)
            Next j
            UserForm1.ReportName.Value = Sheet6.Range("B" & i + 2).Value
            UserForm1.TextBox1.Value = Sheet6.Range("F" & i + 2).Value
            UserForm1.TextBox3.Value = Sheet6.Range("H" & i + 2).Value
            UserForm1.Label10.Caption = Sheet6.Range("A" & i + 2).Value
            UserForm1.ChartPage.Caption = Sheet6.Range("C" & i + 2).Value
            UserForm1.ChartTable.Caption = Sheet6.Range("D" & i + 2).Value
            UserForm1.Label14.Caption = Sheet6.Range("J" & i + 2).Value

            UserForm1.SaveReportAs.Visible = True
            UserForm1.EditReport.Visible = True
            UserForm1.SaveReport.Visible = False
            EditReport
            Dim filters() As String

            If Sheet6.Range("I" & i + 2).Value <> "" Then
                filters = Split(Sheet6.Range("I" & i + 2).Value, ";")
                For j = LBound(filters) To UBound(filters)
                    UserForm1.ListBox3.AddItem filters(j)
                Next j
            End If

            For j = LBound(Fields) To UBound(Fields)
                For k = UserForm1.ListBox1.ListCount - 1 To 1 Step -1
                    If UserForm1.ListBox1.List(k) = Fields(j) Then
                        UserForm1.ListBox1.RemoveItem k
                        Exit For
                    End If
                Next k
            Next j
            With UserForm1
                .Top = Application.Top + Application.Height / 2 - .Height / 2
                .Left = Application.Left + Application.Width / 2 - .Width / 2
            
            End With
            UserForm1.Show
        End If
     End If
Next i

End Sub

Private Sub UserForm_Click()

End Sub
