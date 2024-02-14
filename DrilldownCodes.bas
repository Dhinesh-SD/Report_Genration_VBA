Attribute VB_Name = "DrilldownCodes"
Option Explicit
Sub drillDown(sheet As Worksheet, rng As String, typ As String, Optional Title As String = "DATA TABLE")
Dim returnButton As shape, createReport As shape
Dim ws As Worksheet
Set ws = ActiveSheet
Debug.Print sheet.Range(rng).PivotTable.Parent.name
    On Error GoTo Handler
    'unprotc
    
    sheet.Range(rng).End(xlDown).End(xlToRight).showDetail = True
    ws.Visible = xlSheetHidden
    'protc ws
    Set returnButton = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, Range("A1").Width, Range("A1:A2").Height)
    returnButton.TextFrame.Characters.Text = "BACK TO CHART"
    returnButton.TextFrame.Characters.Font.Size = 14
    returnButton.TextFrame.Characters.Font.Bold = True
    returnButton.TextFrame.VerticalAlignment = xlVAlignCenter
    returnButton.TextFrame.HorizontalAlignment = xlHAlignCenter

    returnButton.ShapeStyle = msoShapeStylePreset36
    
    Set createReport = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Range("D2").Left, Range("D2").Top, Range("D2").Width, Range("D2").Height)
    createReport.TextFrame.Characters.Text = "CREATE CUSTOM REPORT"
    createReport.TextFrame.Characters.Font.Size = 11
    createReport.TextFrame.Characters.Font.Bold = True
    createReport.TextFrame.VerticalAlignment = xlVAlignCenter
    createReport.TextFrame.HorizontalAlignment = xlHAlignCenter
    createReport.ShapeStyle = msoShapeStylePreset72
    createReport.OnAction = "Create_Report"

    ActiveSheet.Range("A1").Value = ""
    ActiveSheet.Range("B1:C2").Merge
    ActiveSheet.Range("B1:C2").Font.Size = 16
    ActiveSheet.Range("B1:C2").Font.Bold = True
    ActiveSheet.Range("B1:C2").Value = Title
    ActiveSheet.Range("D1").Value = "Data from Worksheet Name:"
    ActiveSheet.Range("E1").Value = ws.name
    ActiveSheet.Range("B1").Select
    'https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=1415
Dim i As Long
Dim lastrow As Long
If typ = "Requests" Then
    returnButton.OnAction = "'" & ActiveWorkbook.name & "'!returnToChart"
    lastrow = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ActiveSheet.Cells(i, 1).Value
        If ActiveSheet.Cells(i, 2).Value <> 0 Then ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractDetails.aspx?ID=" & ActiveSheet.Cells(i, 2).Value
    Next i
ElseIf typ = "ContractID" Then
    returnButton.OnAction = "'" & ActiveWorkbook.name & "'!returnToChart"
    lastrow = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractDetails.aspx?ID=" & ActiveSheet.Cells(i, 1).Value
        If ActiveSheet.Cells(i, 2).Value <> 0 Then ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ActiveSheet.Cells(i, 2).Value
    Next i
ElseIf typ = "Tasks" Then
    returnButton.OnAction = "'" & ActiveWorkbook.name & "'!returnToChart"
    lastrow = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
    For i = 4 To lastrow
        ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 1), Address:="https://nedhhs.cobblestone.software/core/ContractRequestTaskDetails.aspx?ID=" & ActiveSheet.Cells(i, 1).Value
        If ActiveSheet.Cells(i, 2).Value <> 0 Then ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(i, 2), Address:="https://nedhhs.cobblestone.software/core/ContractRequestDetails.aspx?ID=" & ActiveSheet.Cells(i, 2).Value
    Next i
End If
If sheet.Range(rng).PivotTable.Parent.name = "Service Req with CLMS num" Then
    For i = 1 To ActiveSheet.Cells(3, 1).End(xlToRight).Column
        ActiveSheet.Cells(3, i).Value = Replace(Replace(ActiveSheet.Cells(3, i).Value, "ContractDetails_with_Names[", ""), "]", "")
        If Trim(ActiveSheet.Cells(3, i).Value) = "Agreement Ends" Or Trim(ActiveSheet.Cells(3, i).Value) = "Agreement Begins" Then
            ActiveSheet.Range(ActiveSheet.ListObjects(1).name & "[" & ActiveSheet.Cells(3, i).Value & "]").NumberFormat = "mm/dd/yyyy"
        End If
    Next i
ElseIf sheet.Range(rng).PivotTable.Parent.name = "TABLES- ALL SERVICE REQ." Or sheet.Range(rng).PivotTable.Parent.name = "TABLE - SERV REQ DATES" Then
    For i = 1 To ActiveSheet.Cells(3, 1).End(xlToRight).Column
        ActiveSheet.Cells(3, i).Value = Replace(Replace(ActiveSheet.Cells(3, i).Value, "Merged_Request_Details[", ""), "]", "")
    Next i
ElseIf sheet.Range(rng).PivotTable.Parent.name = "TABLE-TASK DETAILS" Then
    For i = 1 To ActiveSheet.Cells(3, 1).End(xlToRight).Column
        ActiveSheet.Cells(3, i).Value = Replace(Replace(ActiveSheet.Cells(3, i).Value, "Merge2[", ""), "]", "")
    Next i

End If

Handler:
Debug.Print Err.Number
If Err.Number = 1004 Then
    ws.Select
    MsgBox "Clear All filters and Try again!"
    ws.Shapes("VST").Visible = msoFalse
End If
End Sub

Sub returnToChart()
Application.DisplayAlerts = False
    Dim ws2 As Worksheet
    Dim sheetName As String
    Set ws2 = ThisWorkbook.ActiveSheet
    If Left(ws2.name, 9) <> "DASHBOARD" Then
        sheetName = ws2.Cells(1, 5).Value
        NavigateTo ThisWorkbook.Worksheets(sheetName)
        ws2.Delete
    End If
    If ThisWorkbook.ActiveSheet.name <> Sheet28.name Then Range("A1").Select
Application.DisplayAlerts = True
End Sub
Sub NavigateTo(wsheet As Worksheet)
Application.ScreenUpdating = False
Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets
    If sh.name = wsheet.name Then
        wsheet.Visible = True
        ThisWorkbook.Worksheets(wsheet.name).Select
        Exit For
    End If
Next sh
For Each sh In ThisWorkbook.Worksheets
    'If sh.Name <> wsheet.Name Then sh.Visible = False
Next sh
ActiveWindow.Zoom = 90
Application.ScreenUpdating = True

End Sub
Sub Create_Report()
Dim i As Long, j As Long, k As Long
Dim src As Range, rng As Range, Drng As Range
j = 4
k = 48
Sheet10.Range("D31").CurrentRegion.ClearContents
For i = 1 To ActiveSheet.Cells(3, 1).End(xlToRight).Column
    UserForm1.ListBox1.AddItem ActiveSheet.Cells(3, i).Value
    'Sheet49.Cells(i + 1, 1).Value = ActiveSheet.Cells(3, i).Value
     If ActiveSheet.Range("E1").Value = "DASHBOARD- SERVICE REQUESTS" And ActiveSheet.Cells(3, i).Value = Sheet10.Cells(29, k).Value Then
        Set Drng = Sheet10.Range(Cells(31, k).Address)
        Set rng = ActiveSheet.ListObjects(1).ListColumns(ActiveSheet.Cells(3, i).Value).Range
        rng.AdvancedFilter xlFilterCopy, , Drng, True
        Sheet10.Range(Sheet10.Cells(31, k), Sheet10.Cells(Sheet10.Rows.count, k).End(xlUp)).Offset(1).name = "Filter2_" & k - 3
        k = k + 1
     ElseIf ActiveSheet.Cells(3, i).Value = Sheet10.Cells(29, j).Value Then
        Set Drng = Sheet10.Range(Cells(31, j).Address)
        Set rng = ActiveSheet.ListObjects(1).ListColumns(ActiveSheet.Cells(3, i).Value).Range
        rng.AdvancedFilter xlFilterCopy, , Drng, True
        Sheet10.Range(Sheet10.Cells(31, j), Sheet10.Cells(Sheet10.Rows.count, j).End(xlUp)).Offset(1).name = "Filter_" & j - 3
        j = j + 1
     End If
Next i
On Error Resume Next
Sheet10.Shapes(1).Delete
On Error GoTo 0
UserForm1.TextBox1.Value = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "") & "Custom Reports"
If Sheet6.Range("A2") <> "" Then
    UserForm1.Label10.Caption = Sheet6.Range("A1").End(xlDown).Row
Else
    UserForm1.Label10.Caption = 1
End If
UserForm1.Label11.Caption = "Create New Report:"
UserForm1.ChartPage.Caption = ActiveSheet.Range("E1").Value
UserForm1.ChartTable.Caption = ThisWorkbook.Worksheets(ActiveSheet.Range("E1").Value).Range("A1").Value

With UserForm1
    .Top = Application.Top + Sheet6.Range("A10").Top
    .Left = Application.Left + Sheet6.Range("C10").Left

End With

UserForm1.Show vbModeless

End Sub
Sub EditReport()
Dim i As Long

For i = 1 To ActiveSheet.Cells(3, 1).End(xlToRight).Column
    UserForm1.ListBox1.AddItem ActiveSheet.Cells(3, i).Value
    'Sheet49.Cells(i + 1, 1).Value = ActiveSheet.Cells(3, i).Value
Next i
UserForm1.TextBox1.Value = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "") & "Custom Reports"
'UserForm1.Label10.Caption = Sheet6.Range("A1").End(xlDown).Row - 1

UserForm1.Label11.Caption = "Edit Report":


End Sub
Sub showButton()

Dim tableName As String
Dim VST As shape
Dim ws As Worksheet
Set ws = ActiveSheet
If Left(Application.Caller, 3) = "DS1" Then
    ws.Shapes("VST").Visible = True
    Set VST = ws.Shapes("VST")
    tableName = "Table_" & ActiveSheet.Shapes(Application.Caller).name
    ActiveSheet.Range("A1").Value = tableName
    VST.Left = ActiveSheet.Shapes(Application.Caller).Left + ActiveSheet.Shapes(Application.Caller).Width - VST.Width
    If ActiveSheet.Shapes(Application.Caller).name <> "DS1C3" Then
        VST.Top = ActiveSheet.Shapes(Application.Caller).Top - VST.Height
    Else
        VST.Top = ActiveSheet.Shapes(Application.Caller).Top
    End If
    VST.OnAction = "'" & ActiveWorkbook.name & "'!drillDownButton"
ElseIf Left(Application.Caller, 3) = "DS2" Then
    ws.Shapes("VST").Visible = True
    Set VST = ws.Shapes("VST")
    tableName = "Table_" & ActiveSheet.Shapes(Application.Caller).name
    ActiveSheet.Range("A1").Value = tableName
    VST.Left = ActiveSheet.Shapes(Application.Caller).Left + ActiveSheet.Shapes(Application.Caller).Width - VST.Width
    VST.Top = ActiveSheet.Shapes(Application.Caller).Top
    VST.OnAction = "'" & ActiveWorkbook.name & "'!drillDownButton"
ElseIf Left(Application.Caller, 3) = "DS3" Then
    ws.Shapes("VST").Visible = True
    Set VST = ws.Shapes("VST")
    tableName = "Table_" & ActiveSheet.Shapes(Application.Caller).name
    ActiveSheet.Range("A1").Value = tableName
    VST.Left = ActiveSheet.Shapes(Application.Caller).Left
    VST.Top = ActiveSheet.Shapes(Application.Caller).Top
    VST.OnAction = "'" & ActiveWorkbook.name & "'!drillDownButton"
ElseIf Left(Application.Caller, 3) = "DS4" Then
    ws.Shapes("VST").Visible = True
    Set VST = ws.Shapes("VST")
    tableName = "Table_" & ActiveSheet.Shapes(Application.Caller).name
    ActiveSheet.Range("A1").Value = tableName
    VST.Left = ActiveSheet.Shapes(Application.Caller).Left
    VST.Top = ActiveSheet.Shapes(Application.Caller).Top
    VST.OnAction = "'" & ActiveWorkbook.name & "'!drillDownButton"

Else
    ws.Shapes("VST").Visible = msoFalse
End If
End Sub

Sub refreshPivotTables()
'On Error Resume Next
  Dim PC As PivotTable
  Dim count As Integer
If ActiveSheet.name = Sheet13.name Then
  For Each PC In Sheet9.PivotTables
    PC.PivotCache.Refresh
  Next PC
ElseIf ActiveSheet.name = Sheet15.name Then
  For Each PC In Sheet17.PivotTables
    PC.PivotCache.Refresh
  Next PC

ElseIf ActiveSheet.name = Sheet28.name Then
  For Each PC In Sheet22.PivotTables
    PC.PivotCache.Refresh
  Next PC
ElseIf ActiveSheet.name = Sheet19.name Then
  For Each PC In Sheet43.PivotTables
    PC.PivotCache.Refresh
  Next PC

End If
'unprotc
ActiveSheet.Shapes("LastRefreshed").TextFrame.Characters.Text = "Last Refreshed on: " & Now
'protc
End Sub


Sub drillDownButton()
Application.ScreenUpdating = False

Dim ws As Worksheet
Set ws = ActiveSheet
Dim myChart As ChartObject
Dim Title As String
If ActiveSheet.name = Sheet13.name Then
    Title = ActiveSheet.Shapes("DS1T" & Right(ActiveSheet.Range("A1").Value, 1)).TextFrame.Characters.Text
    drillDown Sheet9, Range("A1").Value, "ContractID", Title
ElseIf ActiveSheet.name = Sheet15.name Then
    Title = ActiveSheet.Shapes("DS2T" & Right(ActiveSheet.Range("A1").Value, 1)).TextFrame.Characters.Text
    drillDown Sheet17, Range("A1").Value, "Requests", Title
ElseIf ActiveSheet.name = Sheet28.name Then
    
    If Left(Right(ActiveSheet.Range("A1").Value, 2), 1) <> "C" Then
        Title = ActiveSheet.Shapes("DS3T" & Right(ActiveSheet.Range("A1").Value, 2)).TextFrame.Characters.Text
    Else
        Title = ActiveSheet.Shapes("DS3T" & Right(ActiveSheet.Range("A1").Value, 1)).TextFrame.Characters.Text
    End If

    'Debug.Print Err.Number
    drillDown Sheet22, Range("A1").Value, "Requests", Title
ElseIf ActiveSheet.name = Sheet19.name Then
    Title = ActiveSheet.Shapes("DS4T" & Right(ActiveSheet.Range("A1").Value, 1)).TextFrame.Characters.Text
    drillDown Sheet43, Range("A1").Value, "Tasks", Title

End If
Application.ScreenUpdating = True
End Sub



Public Sub PivotTable_Collapse()

    Dim pt As PivotTable
    Dim pF As PivotField
    Dim ps As PivotItem

    Set pt = ActiveSheet.PivotTables(1)
    Debug.Print pt.name
    For Each pt In ActiveSheet.PivotTables
        For Each pF In pt.RowFields
            For Each ps In pF.PivotItems
                On Error Resume Next
                ps.DrillTo pF.name
            Next ps
        Next pF
    Next pt

End Sub

