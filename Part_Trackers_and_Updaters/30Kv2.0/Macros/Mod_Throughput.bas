Attribute VB_Name = "Mod_Throughput"
Sub Throughput()


Application.ScreenUpdating = False
On Error Resume Next
Worksheets("NEO 5322121").Activate
LDate = DateAdd("d", -7, Range("A1"))
UDate = Range("A1")
CMonth = Month(Range("A1"))



'Daily
k = 2
For j = 7 To 43
    Count1 = 0
    For Each i In Worksheets("NEO 5322121").Range("C" & j & ":" & "BKJ" & j)
        If i = UDate And i <> "" Then
        Count1 = Count1 + 1
        End If
    Next i
        Worksheets("Throughput").Range("B" & k) = Count1
        k = k + 1
Next j

'Weekly
X = 2
For M = 7 To 43
    Count2 = 0
    For Each n In Worksheets("NEO 5322121").Range("C" & M & ":" & "BKJ" & M)
        If LDate <= n And n <= UDate And n <> "" Then
            Count2 = Count2 + 1
        End If
    Next n
        Worksheets("Throughput").Range("C" & X) = Count2
        X = X + 1
Next M

'Monthly
Y = 2
For f = 7 To 43
    Count3 = 0
    For Each h In Worksheets("NEO 5322121").Range("C" & f & ":" & "BKJ" & f)
        If Month(h) = CMonth And h <> "" Then
            Count3 = Count3 + 1
        End If
    Next h
        Worksheets("Throughput").Range("D" & Y) = Count3
        Y = Y + 1
Next f

Worksheets("Throughput").Activate
End Sub

