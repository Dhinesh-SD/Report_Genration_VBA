VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Create Report"
   ClientHeight    =   11844
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17988
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdB As New Class1
Public enableEvents As Boolean



Private Sub ComboBox1_Click()
If Me.ComboBox1.Value = "Between" Then
    Me.TextBox5.Visible = True
    Me.ComboBox2.Visible = False
    Me.Label13.Visible = True
Else
    Me.TextBox5.Visible = False
    If Me.ComboBox1.RowSource = "Table17[Filters]" Then
        Me.ComboBox2.Visible = True
        Me.TextBox4.Visible = False
    Else
        Me.ComboBox2.Visible = False
        Me.TextBox4.Visible = True
    End If
    Me.Label13.Visible = False

End If
End Sub

Private Sub CommandButton1_Click()
Me.enableEvents = False
   Dim iCtr As Long
   Dim iCtr2 As Long
   Dim isfound As Boolean
   Dim str As String
   
    For iCtr = Me.ListBox1.ListCount - 1 To 0 Step -1
       
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox2.AddItem Me.ListBox1.List(iCtr), 0
            Me.ListBox1.RemoveItem iCtr
        End If
        
    Next iCtr
Me.enableEvents = True
End Sub

Private Sub CommandButton10_Click()
If Right(Me.TextBox2.Value, 13) <> "@nebraska.gov" Then
    Me.Prompt.Caption = "Please Enter a valid Work Email-Id"
    Exit Sub
End If
If Me.TextBox3.Value = "" Then
    Me.TextBox3.Value = Me.TextBox2.Value
Else
    Me.TextBox3.Value = Me.TextBox3.Value & ";" & Me.TextBox2.Value
End If
End Sub

Private Sub CommandButton11_Click()
Unload Me
End Sub

Private Sub CommandButton12_Click()
'Equals, Not Equals, Greater Than, Greater Than or Equal To, Less Than, Less Than or Equal To,
'Between, Contains, Not Contains, Begins With, Ends With
Dim filterList
Dim filter As String
Dim val1, val2 As String
Dim Criteria As String
Criteria = Me.Label12.Caption
If Me.TextBox4.Visible = True Then
    val1 = Me.TextBox4.Value
Else
    val1 = Me.ComboBox2.Value
End If

'If CInt(Me.MultifiltCounter.Caption) = 3 Then
'    MsgBox ("Cant add multiple filters to more than 3 fields!")
'    Exit Sub
'End If

val2 = Me.TextBox5.Value
filter = Me.ComboBox1.Value
    If filter = "Equals" Then
        Me.ListBox3.AddItem Criteria & "|" & "=" & val1
    ElseIf filter = "Not Equals" Then
        Me.ListBox3.AddItem Criteria & "|" & "<>" & val1
    ElseIf filter = "Greater Than" Then
        Me.ListBox3.AddItem Criteria & "|" & ">" & val1
    ElseIf filter = "Greater Than or Equal To" Then
        Me.ListBox3.AddItem Criteria & "|" & ">=" & val1
    ElseIf filter = "Less Than or Equal To" Then
        Me.ListBox3.AddItem Criteria & "|" & "<=" & val1
    ElseIf filter = "Less Than" Then
        Me.ListBox3.AddItem Criteria & "|" & "<" & val1
    ElseIf filter = "Contains" Then
        Me.ListBox3.AddItem Criteria & "|" & "=*" & val1 & "*"
    ElseIf filter = "Not Contains" Then
        Me.ListBox3.AddItem Criteria & "|" & "<>*" & val1 & "*"
    ElseIf filter = "Begins With" Then
        Me.ListBox3.AddItem Criteria & "|" & "=" & val1 & "*"
    ElseIf filter = "Ends With" Then
        Me.ListBox3.AddItem Criteria & "|" & "=*" & val1
    ElseIf filter = "Between" Then
        Me.ListBox3.AddItem Criteria & "|" & ">=" & val1
        Me.ListBox3.AddItem Criteria & "|" & "<=" & val1
    End If
If Me.Label14.Caption = "" Then
    Me.Label14.Caption = Criteria
Else
    Me.Label14.Caption = Me.Label14.Caption & "," & Criteria
End If

Dim count As Integer
Dim distinctFilt() As String
Dim label As String
Dim filtList() As String
Dim i As Long, j As Long
Dim multifiltcount As Integer

multifiltcount = 0
filtList = Split(Me.Label14.Caption, ",")
Dim label2 As String

For i = LBound(filtList) To UBound(filtList)
    If InStr(1, label2, filtList(i)) = 0 Then
        If label2 = "" Then
            label2 = filtList(i)
        Else
            label2 = label2 & "," & filtList(i)
        End If
    End If
Next i

distinctFilt = Split(label2, ",")

For i = LBound(distinctFilt) To UBound(distinctFilt)
    count = 0
    For j = 0 To Me.ListBox3.ListCount - 1
        If Left(Me.ListBox3.List(j), Len(distinctFilt(i))) = distinctFilt(i) Then
        count = count + 1
        End If
    Next j
    If count > 1 Then multifiltcount = multifiltcount + 1
    
    If multifiltcount > 3 Then
        MsgBox ("Cant add multiple filters to more than 3 fields!")
        Me.ListBox3.RemoveItem (Me.ListBox3.ListCount - 1)
        Exit Sub
    End If
Next i

End Sub

Private Sub CommandButton13_Click()
Dim i As Long, mark As Long

For i = Me.ListBox3.ListCount - 1 To 0 Step -1
    If Me.ListBox3.Selected(i) = True Then
        Me.ListBox3.RemoveItem i
        mark = i
        Exit For
    End If
Next i
Dim filters() As String, filtList As String
If i <> mark Then Exit Sub
filters = Split(Me.Label14.Caption, ",")
Me.Label14.Caption = ""
Dim count As Integer

count = 0

For i = LBound(filters) To UBound(filters)
    Debug.Print filters(i), filters(mark)
    If i = mark And count = 0 Then
        count = 1
        GoTo nextLoop
    Else
        If Me.Label14.Caption = "" Then
            Me.Label14.Caption = filters(i)
        Else
            Me.Label14.Caption = Me.Label14.Caption & "," & filters(i)
        End If
    End If
nextLoop:
Next i
End Sub

Private Sub CommandButton2_Click()
Me.enableEvents = False
  Dim iCtr As Long
   Dim iCtr2 As Long
   Dim isfound As Boolean
   Dim str As String
   
    For iCtr = Me.ListBox2.ListCount - 1 To 0 Step -1
        
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
            Me.ListBox2.RemoveItem iCtr
        End If
        
    Next iCtr
Me.enableEvents = True
End Sub

Private Sub CommandButton3_Click()
Dim wb As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet
Set ws1 = ActiveSheet
Dim i As Long, j As Long
Dim addr As String
Dim rng As Range
Dim Crng As Range
Dim Drng As Range
Application.ScreenUpdating = False

Dim filters() As String
Dim filtList() As String
Dim count As Integer
Dim max As Integer
Dim k As Long
Dim distinctFilt() As String
Dim label As String
Dim posi1 As String, pos1() As String
Dim posi2 As String, pos2() As String
Dim posi3 As String, pos3() As String
Dim singleCount As String

filtList = Split(Me.Label14.Caption, ",")
ws1.Cells(3, 1).End(xlToRight).Offset(0, 10).CurrentRegion.ClearContents

For i = LBound(filtList) To UBound(filtList)
    If InStr(1, label, filtList(i)) = 0 Then
        If label = "" Then
            label = filtList(i)
        Else
            label = label & "," & filtList(i)
        End If
    End If
Next i

distinctFilt = Split(label, ",")

For i = LBound(filtList) To UBound(filtList)
    ws1.Cells(3, 1).End(xlToRight).Offset(0, 10 + i).Value = filtList(i)
Next i
Dim multifiltcount As Integer
multifiltcount = 0

For k = LBound(distinctFilt) To UBound(distinctFilt)
    count = 0
    For i = LBound(filtList) To UBound(filtList)
        If distinctFilt(k) = filtList(i) Then
            count = count + 1
            If multifiltcount = 0 And posi1 = "" Then
                posi1 = i
            ElseIf multifiltcount = 1 And posi2 = "" Then
                posi2 = i
            ElseIf multifiltcount = 2 And posi3 = "" Then
                posi3 = i
            End If
            
            If count > 1 And multifiltcount = 0 Then
                posi1 = posi1 & "," & i
            ElseIf count > 1 And multifiltcount = 1 Then
                posi2 = posi2 & "," & i
            ElseIf count > 1 And multifiltcount = 2 Then
                posi3 = posi3 & "," & i
            End If
            
        End If
    Next i
    If count > 1 Then
        multifiltcount = multifiltcount + 1
        pos1 = Split(posi1, ",")
        pos2 = Split(posi2, ",")
        pos3 = Split(posi3, ",")
    Else
        If singleCount = "" Then
        singleCount = distinctFilt(k)
        Else
        singleCount = singleCount & "," & distinctFilt(k)
        End If
    End If
'Max 3 fields that can have more than 1 condition
Next k
Dim l As Long
'Debug.Print posi1, vbNewLine, posi2, vbNewLine, posi3

'Dim filt1() As String
'Dim filt1() As String
'Dim filt1() As String
count = 1

If multifiltcount <> 0 Then
    For i = LBound(pos1) To UBound(pos1)
        For j = LBound(pos2) To UBound(pos2)
            For k = LBound(pos3) To UBound(pos3)
                ReDim filters(1 To 2)
                filters = Split(Me.ListBox3.List(pos3(k)), "|")
                ws1.Cells(3, 1).End(xlToRight).Offset(i + count, 10 + CInt(pos3(k))).Value = "=" & Chr(34) & filters(1) & Chr(34)
                count = count + 1
            Next k
        Next j
        count = count - 1
    Next i
    count = 1
    For i = LBound(pos1) To UBound(pos1)
        For j = LBound(pos2) To UBound(pos2)
            For k = LBound(pos3) To UBound(pos3)
                ReDim filters(1 To 2)
                filters = Split(Me.ListBox3.List(pos2(j)), "|")
                ws1.Cells(3, 1).End(xlToRight).Offset(i + count, 10 + CInt(pos2(j))).Value = "=" & Chr(34) & filters(1) & Chr(34)
                count = count + 1
            Next k
            If UBound(pos3) < 0 Then
                ReDim filters(1 To 2)
                filters = Split(Me.ListBox3.List(pos2(j)), "|")
                ws1.Cells(3, 1).End(xlToRight).Offset(i + count, 10 + CInt(pos2(j))).Value = "=" & Chr(34) & filters(1) & Chr(34)
                count = count + 1
            End If
        Next j
        count = count - 1
    Next i
    count = 1
    For i = LBound(pos1) To UBound(pos1)
        For j = LBound(pos2) To UBound(pos2)
            For k = LBound(pos3) To UBound(pos3)
                ReDim filters(1 To 2)
                filters = Split(Me.ListBox3.List(pos1(i)), "|")
                ws1.Cells(3, 1).End(xlToRight).Offset(i + count, 10 + CInt(pos1(i))).Value = "=" & Chr(34) & filters(1) & Chr(34)
                count = count + 1
            Next k
        Next j
        If UBound(pos2) < 0 Then
            ReDim filters(1 To 2)
            filters = Split(Me.ListBox3.List(pos1(i)), "|")
            ws1.Cells(3, 1).End(xlToRight).Offset(i + count, 10 + CInt(pos1(i))).Value = "=" & Chr(34) & filters(1) & Chr(34)
            count = count + 1
        End If
    
        count = count - 1
    Next i
End If
Dim Lrow As Long
Lrow = ws1.Cells(3, 1).End(xlToRight).Offset(0, 10).CurrentRegion.Rows.count - 1

Dim singlefilt() As String
singlefilt = Split(singleCount, ",")
If Lrow = 0 Then Lrow = 1
For i = LBound(singlefilt) To UBound(singlefilt)
    For j = 0 To Me.ListBox3.ListCount - 1
    Debug.Print Left(Me.ListBox3.List(j), Len(singlefilt(i))), singlefilt(i)
        If Trim(LCase(Left(Me.ListBox3.List(j), Len(singlefilt(i))))) = Trim(LCase(singlefilt(i))) Then
            ReDim filters(1 To 2)
            filters = Split(Me.ListBox3.List(j), "|")
            For k = 1 To Lrow
                ws1.Cells(3, 1).End(xlToRight).Offset(k, 10 + j).Value = "=" & Chr(34) & filters(1) & Chr(34)
            Next k
        End If
    Next j
    
Next i
Set wb = Workbooks.Add

Set ws2 = wb.Sheets(1)

For i = 0 To Me.ListBox2.ListCount - 1
    'ws2.Cells(3, i + 1).Value = Me.ListBox2.List(i)
Next i

addr = ws1.Range("A3").End(xlToRight).Address
If ws1.Range("A3").End(xlToRight).Column <= 26 Then
    addr = Left(addr, 3)
Else
    addr = Left(addr, 4)
End If

Set rng = ws1.Range("A3:" & addr & ws1.Range("A3").End(xlDown).Row)
Set Crng = ws1.Cells(3, 1).End(xlToRight).Offset(0, 10).CurrentRegion
Set Drng = ws2.Range("A3:" & addr & "3")


rng.AdvancedFilter xlFilterCopy, Crng, Drng
Dim lastrow As Long
Dim lastcol As Long

If ws1.Range("E1").Value = "DASHBOARD- SERVICE REQUESTS" Then
    lastrow = ws2.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractDetails.aspx?ID=" & ws2.Cells(i, 1).Value
        If ws2.Cells(i, 2).Value <> 0 Then ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ws2.Cells(i, 2).Value
    Next i
ElseIf ws1.Range("E1").Value = "DASHBOARD - TASK DETAILS" Then
    lastrow = ws2.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractRequestTaskDetails.aspx?ID=" & ws2.Cells(i, 1).Value
        If ws2.Cells(i, 2).Value <> 0 Then ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ws2.Cells(i, 2).Value
    Next i

Else
    lastrow = ws2.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ws2.Cells(i, 1).Value
        If ws2.Cells(i, 2).Value <> 0 Then ws2.Hyperlinks.Add Anchor:=ws2.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractDetails.aspx?ID=" & ws2.Cells(i, 2).Value
    Next i

End If


Dim objTable As ListObject
Dim src As Range
Dim isfound As Boolean

lastcol = ws2.Cells(3, 1).End(xlToRight).Column


For j = lastcol To 1 Step -1
    isfound = False
    For i = 0 To Me.ListBox2.ListCount - 1
        If ws2.Cells(3, j).Value = Me.ListBox2.List(i) Then
            isfound = True

            Exit For
        End If
    Next i
    If isfound = False Then
        ws2.Cells(1, j).EntireColumn.Delete
        lastcol = lastcol - 1
    End If

Next j

Set src = ws2.Range("A3").CurrentRegion
Set objTable = ws2.ListObjects.Add(xlSrcRange, src, , xlYes)
On Error Resume Next
ws2.Shapes(1).Delete
On Error GoTo 0
ws2.Range("A1:D1").Merge
ws2.Columns.AutoFit
ws2.Range("A1").Value = ws1.Range("B1").Value & "-" & Now
Dim name As String
name = Me.ReportName.Value
If name = "" Then name = "Report_" & Sheet6.Range("A1").End(xlDown).Row
If Right(Me.TextBox1.Value, 1) = "\" Then
    wb.SaveAs Me.TextBox1.Value & name & "_" & Month(Now) & Day(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xlsx"
Else
    wb.SaveAs Me.TextBox1.Value & "\" & name & "_" & Month(Now) & Day(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xlsx"

End If
Dim EmailList() As String

If Me.TextBox3.Value <> "" Then
EmailList = Split(Me.TextBox3.Value, ";")
For i = LBound(EmailList) To UBound(EmailList)

    Dim outApp As Object, outmail As Object
    Set outApp = CreateObject("Outlook.application")
    Set outmail = outApp.CreateItem(0)
    With outmail
        .To = EmailList(i)
        .CC = ""
        .Subject = name & ":Report Generated From " & ThisWorkbook.name
        .Body = "This Report Was Generated On-" & Now
        .Attachments.Add wb.FullName
        '.Display
        .send
    End With
Next i
End If
'Dim AnswerYesNo As String
'
'AnswerYesNo = MsgBox(" Do you want to view the Report Now?", vbYesNo, "View Report?")

'If AnswerYesNo = vbNo Then
    wb.Close True
    Unload Me
    On Error Resume Next
    ws1.Activate
    If Me.Label11.Caption = "Create New Report:" Then returnToChart
    ActiveWindow.ScrollColumn = 1
'Else
'    Unload Me
'    ws1.Activate
'    returnToChart
'    ActiveWindow.ScrollColumn = 1
'    ws2.Activate
'End If

Dim DataObj As New MSForms.DataObject 'empty the clipboard
DataObj.SetText ""
DataObj.PutInClipboard

Application.ScreenUpdating = True

End Sub


Private Sub CommandButton9_Click()
Dim FldrPicker As FileDialog
Dim myFolder As String

'Have User Select Folder to Save to with Dialog Box
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

  With FldrPicker
    .Title = "Select A Target Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
    myFolder = .SelectedItems(1) & "\"
  End With
  
'Carry out rest of your code here....
Me.TextBox1.Value = myFolder
End Sub

Private Sub EditReport_Click()

Dim name As String
Dim Fields As String
name = Me.ReportName.Value
If name = "" Then name = "Report_" & Sheet6.Range("A1").End(xlDown).Row - 1

Dim lastrow As Long
Dim lastcol As Long
Dim i As Long
Fields = Me.ListBox2.List(0)
For i = 1 To Me.ListBox2.ListCount - 1
    Fields = Fields & "," & Me.ListBox2.List(i)
Next i
lastcol = Sheet6.Range("A1").End(xlToRight).Column
lastrow = CInt(Me.Label10.Caption)

Sheet6.Range("A" & lastrow + 1).Value = lastrow
Sheet6.Range("B" & lastrow + 1).Value = name
Sheet6.Range("C" & lastrow + 1).Value = ActiveSheet.Range("E1").Value
Sheet6.Range("D" & lastrow + 1).Value = ThisWorkbook.Worksheets(ThisWorkbook.ActiveSheet.Range("E1").Value).Range("A1").Value
Sheet6.Range("E" & lastrow + 1).Value = Fields
Sheet6.Range("F" & lastrow + 1).Value = Me.TextBox1.Value
Sheet6.Range("G" & lastrow + 1).Value = Me.CheckBox1.Value
Sheet6.Range("H" & lastrow + 1).Value = Me.TextBox3.Value
Sheet6.Range("H" & lastrow + 1).Value = Me.TextBox3.Value
Sheet6.Range("j" & lastrow + 1).Value = Me.Label14.Caption

Dim filters As String
For i = 0 To Me.ListBox3.ListCount - 1
    If filters = "" Then
        filters = Me.ListBox3.List(i)
    Else
        filters = filters & ";" & Me.ListBox3.List(i)
    End If
Next i
Sheet6.Range("I" & lastrow + 1).Value = filters


Me.Prompt.Caption = "Report Saved Successfully as : " & name
End Sub

Private Sub Label1_Click()

End Sub

Private Sub ListBox2_Change()
Dim i As Long, j As Long
If Me.enableEvents = False Then
    Exit Sub
End If
For i = 0 To Me.ListBox2.ListCount - 1
    If Me.ListBox2.Selected(i) = True Then
        Me.Label12.Caption = Me.ListBox2.List(i)
    End If
Next i

Dim fieldname As String
Dim notFound As Boolean
Dim rsource As String

notFound = True
fieldname = Me.Label12.Caption
If ActiveSheet.Range("E1").Value = "DASHBOARD- SERVICE REQUESTS" Then
    i = 48
Else
    i = 4
End If

For j = i To Sheet10.Range("CZ29").End(xlToLeft).Column
    If fieldname = Sheet10.Cells(29, j).Value Then
        notFound = False
        If (j >= 4 And j < 27) Then
            Me.ComboBox2.Visible = True
            Me.ComboBox2.RowSource = "Filter_" & j - 3
            Me.TextBox4.Visible = False
        ElseIf (j >= 48 And j <= 68) Then
            Me.ComboBox2.Visible = True
            Me.ComboBox2.RowSource = "Filter2_" & j - 3
            Me.TextBox4.Visible = False
        Else
            Me.ComboBox2.Visible = False
            Me.TextBox4.Visible = True
        End If
        If (j >= 27 And j <= 44) Or (j > 69 And j <= 81) Then
            Me.ComboBox1.RowSource = "Table19[Filters]"
        ElseIf (j >= 4 And j <= 26) Or (j >= 48 And j < 69) Then
            Me.ComboBox1.RowSource = "Table17[Filters]"
        End If
        Exit For
    End If
Next j

If notFound = True Then
    Me.ComboBox1.RowSource = "Table20[Filters]"
    Me.ComboBox2.Visible = False
    Me.TextBox4.Visible = True

End If
Me.ComboBox1.Value = Sheet10.Cells(Range(Me.ComboBox1.RowSource).Row, Range(Me.ComboBox1.RowSource).Column).Value
Me.ComboBox2.Value = ""
Me.TextBox4.Value = ""
Me.TextBox5.Value = ""
End Sub


Private Sub MultifiltCounter_Click()

End Sub

Private Sub SaveReport_Click()
Dim name As String
Dim Fields As String
name = Me.ReportName.Value
If name = "" Then name = "Report_" & Me.Label10.Caption

Dim lastrow As Long
Dim lastcol As Long
Dim i As Long
Fields = Me.ListBox2.List(0)
For i = 1 To Me.ListBox2.ListCount - 1
    Fields = Fields & "," & Me.ListBox2.List(i)
Next i
lastcol = Sheet6.Range("A1").End(xlToRight).Column
lastrow = Sheet6.Range("A1").End(xlDown).Row
If Sheet6.Cells(lastrow, 1).Value = "" Then
    lastrow = lastrow - 1
    Sheet6.Range("A" & lastrow + 1).Value = lastrow
Else
    Sheet6.Range("A" & lastrow + 1).Value = Me.Label10.Caption
End If
Sheet6.Range("B" & lastrow + 1).Value = name
Sheet6.Range("C" & lastrow + 1).Value = Me.ChartPage.Caption
Sheet6.Range("D" & lastrow + 1).Value = Me.ChartTable.Caption
Sheet6.Range("E" & lastrow + 1).Value = Fields
Sheet6.Range("F" & lastrow + 1).Value = Me.TextBox1.Value
Sheet6.Range("G" & lastrow + 1).Value = Me.CheckBox1.Value
Sheet6.Range("H" & lastrow + 1).Value = Me.TextBox3.Value
Sheet6.Range("j" & lastrow + 1).Value = Me.Label14.Value

Dim filters As String
For i = 0 To Me.ListBox3.ListCount - 1
    If filters = "" Then
        filters = Me.ListBox3.List(i)
    Else
        filters = filters & ";" & Me.ListBox3.List(i)
    End If
Next i
Sheet6.Range("I" & lastrow + 1).Value = filters

Me.Prompt.Caption = "Report Saved Successfully as : " & name
End Sub


Private Sub SaveReportAs_Click()
Dim name As String
Dim Fields As String
name = Me.ReportName.Value
If name = "" Then name = "Report_" & Sheet6.Range("A1").End(xlDown).Row - 1

Dim lastrow As Long
Dim lastcol As Long
Dim i As Long
Fields = Me.ListBox2.List(0)
For i = 1 To Me.ListBox2.ListCount - 1
    Fields = Fields & "," & Me.ListBox2.List(i)
Next i
lastcol = Sheet6.Range("A1").End(xlToRight).Column
lastrow = Sheet6.Range("A1").End(xlDown).Row
name = name & "_" & lastrow - 1
Sheet6.Range("A" & lastrow + 1).Value = lastrow - 1
Sheet6.Range("B" & lastrow + 1).Value = name
Sheet6.Range("C" & lastrow + 1).Value = Me.ChartPage.Caption
Sheet6.Range("D" & lastrow + 1).Value = Me.ChartTable.Caption
Sheet6.Range("E" & lastrow + 1).Value = Fields
Sheet6.Range("F" & lastrow + 1).Value = Me.TextBox1.Value
Sheet6.Range("G" & lastrow + 1).Value = Me.CheckBox1.Value
Sheet6.Range("H" & lastrow + 1).Value = Me.TextBox3.Value
Sheet6.Range("H" & lastrow + 1).Value = Me.TextBox3.Value
Sheet6.Range("j" & lastrow + 1).Value = Me.Label14.Caption

Dim filters As String
For i = 0 To Me.ListBox3.ListCount - 1
    If filters = "" Then
        filters = Me.ListBox3.List(i)
    Else
        filters = filters & ";" & Me.ListBox3.List(i)
    End If
Next i
Sheet6.Range("I" & lastrow + 1).Value = filters


Me.Prompt.Caption = "Report Saved Successfully as : " & name
End Sub


Private Sub UserForm_Deactivate()
Sheet49.Range(Sheet49.ListObjects(1).name).ClearContents
Sheet49.ListObjects(1).Resize (Range("A1:A2"))
returnToChart
End Sub

Private Sub UserForm_Initialize()
Me.enableEvents = True
End Sub

Private Sub UserForm_Terminate()
'Sheet49.Range(Sheet49.ListObjects(1).name).ClearContents
'Sheet49.ListObjects(1).Resize (Sheet49.Range("A1:A2"))
returnToChart
End Sub
