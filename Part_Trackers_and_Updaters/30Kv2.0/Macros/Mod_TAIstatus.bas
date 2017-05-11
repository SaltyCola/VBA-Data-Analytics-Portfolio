Attribute VB_Name = "Mod_TAIstatus"
Sub TAIstatus()


Application.ScreenUpdating = False
On Error Resume Next
Worksheets("NEO 5322121").Activate
LDate = DateAdd("d", -7, Range("A1"))
UDate = Range("A1")
CMonth = Month(Range("A1"))


Count1 = 0
    For Each i In Worksheets("NEO 5322121").Range("C6:BKJ6")
    If i.Offset(24, 0) = "" And i.Offset(25, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(26, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(27, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(32, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(33, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(34, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(35, 0) <> "" Then
        Count1 = Count1 + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(36, 0) <> "" Then
        Count1 = Count1 + 1
       End If
    Next i

Count2 = 0
j = 4
    For Each i In Worksheets("NEO 5322121").Range("C6:BKJ6")
    If i.Offset(24, 0) <> "" And i.Offset(23, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
       ElseIf i.Offset(24, 0) = "" And i.Offset(22, 0) <> "" Then
       Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(21, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(20, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(19, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(18, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(17, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(16, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(15, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(14, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(13, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(12, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(11, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(10, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(9, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(8, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(7, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(6, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(5, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(4, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
        ElseIf i.Offset(24, 0) = "" And i.Offset(3, 0) <> "" Then
        Count2 = Count2 + 1
        Worksheets("TAI Status").Range("R" & j) = i.Value
        j = j + 1
       End If
    Next i

Worksheets("TAI Status").Range("B4") = Count1
Worksheets("TAI Status").Range("C4") = Count2
Worksheets("TAI Status").Activate



'    lastrow = Application.CountA(Worksheets("TAI Status").Range("R4:A10000")) + 3
'    Range("E" & lastrow3).Select
'    ActiveCell.FormulaR1C1 = _
'        "=IFERROR(INDEX(Database!C[-3],MATCH('Tray Sequence List'!C[-4],Database!C[-4],0)),"" "")"
'    Range("E" & lastrow3).Select
'    Selection.AutoFill Destination:=Range("E" & lastrow3 & ":" & "E" & lastrow4)






      
End Sub


