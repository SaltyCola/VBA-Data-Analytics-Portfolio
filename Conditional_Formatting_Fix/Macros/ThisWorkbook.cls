VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetActivate(ByVal Sh As Object)

    Dim StartLine As Long
    Dim i As Double
    
    'sheet name has "Set"
    If (Left(Sh.Name, 3) = "Set") And (ActiveWorkbook.VBProject.VBComponents(Sh.CodeName).CodeModule.Lines(1, 1) = "") Then
        
        With ActiveWorkbook.VBProject.VBComponents(Sh.CodeName).CodeModule
            StartLine = .CreateEventProc("Change", "Worksheet") + 1
            i = 3
            .InsertLines i, "    'Unrecognized action"
            i = i + 1
            .InsertLines i, "    On Error GoTo lineExitSub"
            i = i + 1
            .InsertLines i, ""
            i = i + 1
            .InsertLines i, "    'Column E"
            i = i + 1
            .InsertLines i, "    If Not Intersect(Target, Range(" & Chr(34) & "E:E" & Chr(34) & ")) Is Nothing Then"
            i = i + 1
            .InsertLines i, "        Select Case Target"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(3, 1).Value: Call GreenCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(4, 1).Value: Call LightGreenCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(5, 1).Value: Call WhiteCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(6, 1).Value: Call OrangeCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(7, 1).Value: Call YellowCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(8, 1).Value: Call RedCell"
            i = i + 1
            .InsertLines i, "            On Error Resume Next"
            i = i + 1
            .InsertLines i, "        End Select"
            i = i + 1
            .InsertLines i, "    End If"
            i = i + 1
            .InsertLines i, ""
            i = i + 1
            .InsertLines i, "    'Column F"
            i = i + 1
            .InsertLines i, "    If Not Intersect(Target, Range(" & Chr(34) & "F:F" & Chr(34) & ")) Is Nothing Then"
            i = i + 1
            .InsertLines i, "        Select Case Target"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(3, 2).Value: Call WhiteCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(4, 2).Value: Call LightYellowCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(5, 2).Value: Call YellowCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(6, 2).Value: Call DarkYellowCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(7, 2).Value: Call BlueCell"
            i = i + 1
            .InsertLines i, "            Case Worksheets(" & Chr(34) & "List Source" & Chr(34) & ").Cells(8, 2).Value: Call RedCell"
            i = i + 1
            .InsertLines i, "            On Error Resume Next"
            i = i + 1
            .InsertLines i, "        End Select"
            i = i + 1
            .InsertLines i, "    End If"
            i = i + 1
            .InsertLines i, ""
            i = i + 1
            .InsertLines i, "lineExitSub:"
            i = i + 1
            .InsertLines i, "    Exit Sub"
            i = i + 1
            .InsertLines i, ""
        End With
    
    End If

End Sub

