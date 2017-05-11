Attribute VB_Name = "Mod_Summary"
Public CellCurrent As Range
Public CellAbove As Range
Public CellsClearAbove As Boolean
Public CellsAboveRange38 As Range
Public CellsAboveRange7 As Range
Public cllsabvRng As Range
Public NEOwrks As Worksheet

'===================================30K SUMMARY PROGRAM=================================='
'clear out all values and colors
Sub Clear()
    'Color Variables
    Dim clrBlnk As Long
    Dim txtBlack As Long

    'Set Color Variables
    clrBlnk = Worksheets("Tracker Summaries").Range("B4").Interior.Color
    txtBlack = Worksheets("Tracker Summaries").Range("A4").Font.Color

    'Clear Cells
    Range("D6:AN313").ClearContents
    Range("D6:AN313").Cells.Interior.Color = clrBlnk
    Range("D6:AN313").Cells.Font.Color = txtBlack
    Range("D6:AN313").Cells.Font.Bold = False

End Sub

Sub Main()
    
    'Color Variables
    Dim clrBlnk As Long
    Dim clrBlack As Long
    Dim clrRed As Long
    Dim clrOrange As Long
    Dim clrPink As Long
    Dim clrBlue As Long
    Dim clrGreen As Long
    Dim clrbrtGreen As Long
    Dim clrPurple As Long
    Dim clrdrkGreen As Long
    Dim clrlghtGreen1 As Long
    Dim clrlghtGreen2 As Long
    Dim clrYellow As Long
    Dim txtYellow As Long
    Dim txtBlack As Long
    'Other Variables
    Dim sData As Range
    Dim sInt As Range
    Dim sRow As Long
    Dim sCol As Long
    Dim sStart As Long
    Dim sEnd As Long
    Dim rData As Range
    Dim rStart As Range
    Dim rEnd As Range
    Dim summData As Range
    Dim summCol As Long
    Dim colData As Range
    Dim colInt As Range
    Dim rNum As Long
    Dim fData As Range
    Dim fInt As Range
    Dim fRow As Long
    Dim rw As Range
    
    
    'Set Color Variables
    clrBlnk = Worksheets("Tracker Summaries").Range("B4").Interior.Color
    clrBlack = Worksheets("Tracker Summaries").Range("B15").Interior.Color
    clrRed = Worksheets("Tracker Summaries").Range("C6").Interior.Color
    clrOrange = Worksheets("Tracker Summaries").Range("B6").Interior.Color
    clrPink = Worksheets("Tracker Summaries").Range("B7").Interior.Color
    clrBlue = Worksheets("Tracker Summaries").Range("B8").Interior.Color
    clrGreen = Worksheets("Tracker Summaries").Range("B9").Interior.Color
    clrbrtGreen = Worksheets("Tracker Summaries").Range("B16").Interior.Color
    clrPurple = Worksheets("Tracker Summaries").Range("B17").Interior.Color
    clrdrkGreen = Worksheets("Tracker Summaries").Range("B10").Interior.Color
    clrlghtGreen1 = Worksheets("Tracker Summaries").Range("B11").Interior.Color
    clrlghtGreen2 = Worksheets("Tracker Summaries").Range("B12").Interior.Color
    clrYellow = Worksheets("Tracker Summaries").Range("B14").Interior.Color
    txtYellow = Worksheets("Tracker Summaries").Range("B13").Font.Color
    txtBlack = Worksheets("Tracker Summaries").Range("A4").Font.Color
    
    
    'Clear Cells
    Range("D6:AN313").ClearContents
    Range("D6:AN313").Cells.Interior.Color = clrBlnk
    Range("D6:AN313").Cells.Font.Color = txtBlack
    Range("D6:AN313").Cells.Font.Bold = False
    
    
    'Turn off ScreenUpdating
    Application.ScreenUpdating = False
    
    
    'Unhide Rows
    For Each rw In Range("D6:D313")
        rw.EntireRow.Hidden = False
    Next rw
    
    
    'Find rData Length
    Set sData = Worksheets("NEO 5322121").Range("6:6")
    sStart = 0
    sEnd = 0
    sRow = 0
    sCol = 0
    'Find starting point
    For Each sInt In sData
        If (sInt.EntireColumn.Hidden = False) And (sInt.Column > 2) Then
            sStart = sInt.Column
            Exit For
        End If
    Next sInt
    'Find ending point
    For Each sInt In sData
        sRow = sInt.Row
        sCol = sInt.Column
        If (sInt.EntireColumn.Hidden = False) And (IsEmpty(sInt) = False) And (Worksheets("NEO 5322121").Cells(sRow, (sCol + 1)).Interior.Color = RGB(255, 0, 0)) Then
            sEnd = sInt.Column
            Exit For
        End If
    Next sInt
    
    
    'Turn on ScreenUpdating
    Application.ScreenUpdating = True
    
    
    'Iterate Summary Sheet Columns
    Set colData = Worksheets("Tracker Summaries").Range("D1:AN1")
    For Each colInt In colData
        rNum = colInt.Value
        summCol = colInt.Offset(1, 0).Value
        
        'Set Ranges
        Set rStart = Worksheets("NEO 5322121").Cells(rNum, sStart)
        Set rEnd = Worksheets("NEO 5322121").Cells(rNum, sEnd)
        Set rData = Worksheets("NEO 5322121").Range(rStart, rEnd)
        Set summData = Worksheets("Tracker Summaries").Range(Cells(14, summCol), Cells(313, summCol))
        
        'Count Totals Sub
        Call CountSnOpLocations(sEnd, rData, summData, summCol, clrBlnk, clrBlack, clrRed, clrOrange, clrPink, clrBlue, clrGreen, clrbrtGreen, clrPurple, clrdrkGreen, clrlghtGreen1, clrlghtGreen2, clrYellow, txtYellow)
        'Grab SN Info Sub
        Call GrabSnInfo(sEnd, rData, summData, summCol, clrBlnk, clrBlack, clrRed, clrOrange, clrPink, clrBlue, clrGreen, clrbrtGreen, clrPurple, clrdrkGreen, clrlghtGreen1, clrlghtGreen2, clrYellow, txtYellow)
    
    Next colInt
    
    
    'Turn off ScreenUpdating
    Application.ScreenUpdating = False
    
    
    'Hide Unused Rows in Summary
    Set fData = Worksheets("Tracker Summaries").Range("D14", "D311")
    fRow = 0
    For Each fInt In fData
        fRow = fInt.Row
        If Application.CountA(fInt.EntireRow) = 0 Then
            Rows(fRow).Hidden = True
        End If
    Next fInt
    
    
    'Turn on ScreenUpdating
    Application.ScreenUpdating = True
    
    
End Sub


'Search and Count filled cells with blank cell directly above
Sub CountSnOpLocations(ByVal sEnd As Long, ByVal rData As Range, ByVal summData As Range, ByVal summCol As Integer, ByVal clrBlnk As Long, ByVal clrBlack As Long, ByVal clrRed As Long, ByVal clrOrange As Long, ByVal clrPink As Long, ByVal clrBlue As Long, ByVal clrGreen As Long, ByVal clrbrtGreen As Long, ByVal clrPurple As Long, ByVal clrdrkGreen As Long, ByVal clrlghtGreen1 As Long, ByVal clrlghtGreen2 As Long, ByVal clrYellow As Long, ByVal txtYellow As Long)
    
    Dim rInt As Range
    Dim rowNum As Integer
    Dim colNum As Integer
    Dim clrRefCell As Long
    'Color Counts
    Dim cntRed As Long
    Dim cntOrange As Long
    Dim cntPink As Long
    Dim cntPinkdrk As Long
    Dim cntPinklght As Long
    Dim cntPinkPurple As Long
    Dim cntBlue As Long
    Dim cntGreen As Long
    Dim cntPurple As Long
    Dim cntdrkGreen As Long
    Dim cntlghtGreen As Long
    'Totals cells info
    Dim badCellColor As Long
    Dim badCellVal As Integer
    Dim slowCellColor As Long
    Dim slowCellVal As Integer
    Dim rtoCellColor As Long
    Dim rtoCellVal As Integer
    Dim goodCellColor As Long
    Dim goodCellVal As Integer

    rowNum = rData.Row
    clrRefCell = Worksheets("Tracker Summaries").Cells(3, summCol).Interior.Color
    cntRed = 0
    cntOrange = 0
    cntPink = 0
    cntPinkdrk = 0
    cntPinklght = 0
    cntPinkPurple = 0
    cntBlue = 0
    cntGreen = 0
    cntPurple = 0
    cntdrkGreen = 0
    cntlghtGreen = 0
    
    'Iterate tracker sheet
    For Each rInt In rData
        colNum = rInt.Column
        
        'set public variables
        Set NEOwrks = ActiveWorkbook.Worksheets("NEO 5322121")
        CellsClearAbove = False
        Set CellCurrent = rInt
        If CellCurrent.Row = 38 Then
           Set CellAbove = CellCurrent.Offset(-5, 0)
        ElseIf CellCurrent.Row = 7 Then
            Set CellAbove = CellCurrent
        Else
            Set CellAbove = CellCurrent.Offset(-1, 0)
        End If
        
        'set cells clear above range
        If CellAbove.Row > 38 Then
            Application.ScreenUpdating = False
            NEOwrks.Activate
            Set CellsAboveRange38 = Range(Cells(CellAbove.Row, CellAbove.Column), Cells(38, colNum))
            Set CellsAboveRange7 = Range(Cells(33, colNum), Cells(7, colNum))
            Worksheets("Tracker Summaries").Activate
            Application.ScreenUpdating = True
        ElseIf CellAbove.Row <= 33 Then
            Application.ScreenUpdating = False
            NEOwrks.Activate
            Set CellsAboveRange38 = NEOwrks.Range(Cells(CellAbove.Row, CellAbove.Column), Cells(7, colNum))
            Set CellsAboveRange7 = NEOwrks.Range(Cells(CellAbove.Row, CellAbove.Column), Cells(7, colNum))
            Worksheets("Tracker Summaries").Activate
            Application.ScreenUpdating = True
        End If
        'set cells clear above boolean
        CellsClearAbove = True
        For Each cllsabvRng In CellsAboveRange38
            If Not cllsabvRng.Interior.Color = clrBlnk Then
                CellsClearAbove = False
                Exit For
            End If
        Next cllsabvRng
        For Each cllsabvRng In CellsAboveRange7
            If Not cllsabvRng.Interior.Color = clrBlnk Then
                CellsClearAbove = False
                Exit For
            End If
        Next cllsabvRng
        
        'Fill hidden rows 34 through 37 on Tracker
        If Worksheets("NEO 5322121").Cells(33, colNum).Interior.Color = clrGreen Then
            Worksheets("NEO 5322121").Cells(34, colNum).Interior.Color = clrGreen
            Worksheets("NEO 5322121").Cells(35, colNum).Interior.Color = clrGreen
            Worksheets("NEO 5322121").Cells(36, colNum).Interior.Color = clrGreen
            Worksheets("NEO 5322121").Cells(37, colNum).Interior.Color = clrGreen
        End If
        
    'Find Non-Pink colored cell with conditions
        If (rInt.EntireRow.Hidden = False) And (rInt.EntireColumn.Hidden = False) And Not (clrPink = Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color) And Not (clrBlnk = rInt.Interior.Color) And CellsClearAbove And Not (clrBlack = rInt.Interior.Color) Then
            'Color Counter
            If rInt.Interior.Color = clrRed Then
                cntRed = cntRed + 1
            End If
            If rInt.Interior.Color = clrOrange Then
                cntOrange = cntOrange + 1
            End If
            If rInt.Interior.Color = clrBlue Then
                cntBlue = cntBlue + 1
            End If
            If rInt.Interior.Color = clrGreen Then
                cntGreen = cntGreen + 1
            End If
            'bright greens (really only for launch column)============
            If rInt.Interior.Color = clrbrtGreen Then
                cntGreen = cntGreen + 1
            End If
            '=========================================================
            If rInt.Interior.Color = clrPurple Then
                cntPurple = cntPurple + 1
            End If
            If rInt.Interior.Color = clrdrkGreen Then
                cntdrkGreen = cntdrkGreen + 1
            End If
            If rInt.Interior.Color = clrlghtGreen1 Then
                cntlghtGreen = cntlghtGreen + 1
            End If
            If rInt.Interior.Color = clrlghtGreen2 Then
                cntlghtGreen = cntlghtGreen + 1
            End If
        End If
        
    'Find Pink Cells with conditions
        If (rInt.EntireRow.Hidden = False) And (rInt.EntireColumn.Hidden = False) And (clrPink = Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color) And Not (clrBlnk = rInt.Interior.Color) And CellsClearAbove And Not (clrBlack = rInt.Interior.Color) Then
            'normal green
            If Not (rInt.Interior.Color = clrdrkGreen) Then
                cntPink = cntPink + 1
            End If
            'purple (in house from outsourced)
            If rInt.Interior.Color = clrPurple Then
                cntPinkPurple = cntPinkPurple + 1
            End If
            'dark green
            If rInt.Interior.Color = clrdrkGreen Then
                cntPinkdrk = cntPinkdrk + 1
            End If
            'light green
            If rInt.Interior.Color = clrlghtGreen1 Then
                cntPinklght = cntPinklght + 1
            End If
            If rInt.Interior.Color = clrlghtGreen2 Then
                cntPinklght = cntPinklght + 1
            End If
        End If
        
    Next rInt
    
    'Totals cell colors
    badCellColor = clrOrange
    slowCellColor = clrPink
    rtoCellColor = clrBlue
    goodCellColor = clrGreen
    'Totals cell values
    badCellVal = cntOrange
    slowCellVal = cntPink
    rtoCellVal = cntBlue
    goodCellVal = cntGreen
    
    'Alternate Totals Colors
    If clrRefCell = clrRed Then
        badCellVal = cntRed + cntOrange
    End If
    If clrRefCell = clrPurple Then
        goodCellVal = cntPurple + cntGreen
        slowCellVal = cntPinkPurple + cntPink
    End If
    If clrRefCell = clrdrkGreen Then
        goodCellColor = clrdrkGreen
        goodCellVal = cntdrkGreen
        'remove other color counts
        badCellVal = 0
        slowCellVal = cntPinkdrk
        rtoCellVal = 0
    End If
    If clrRefCell = RGB(175, 255, 175) Then
        goodCellVal = cntdrkGreen + cntGreen
        slowCellVal = cntPinkdrk + cntPink
    End If
    If clrRefCell = clrlghtGreen1 Then
        goodCellColor = clrGreen
        goodCellVal = cntGreen + cntlghtGreen
        slowCellVal = cntPink + cntPinklght
    End If
    
    'Zero Value No Fill
    If badCellVal = 0 Then
        badCellColor = clrBlnk
    End If
    If slowCellVal = 0 Then
        slowCellColor = clrBlnk
    End If
    If rtoCellVal = 0 Then
        rtoCellColor = clrBlnk
    End If
    If goodCellVal = 0 Then
        goodCellColor = clrBlnk
    End If
    
    'Fill Totals Cells
    Worksheets("Tracker Summaries").Cells(6, summCol).Value = badCellVal
    Worksheets("Tracker Summaries").Cells(6, summCol).Interior.Color = badCellColor
    Worksheets("Tracker Summaries").Cells(7, summCol).Value = slowCellVal
    Worksheets("Tracker Summaries").Cells(7, summCol).Interior.Color = slowCellColor
    Worksheets("Tracker Summaries").Cells(8, summCol).Value = rtoCellVal
    Worksheets("Tracker Summaries").Cells(8, summCol).Interior.Color = rtoCellColor
    Worksheets("Tracker Summaries").Cells(9, summCol).Value = goodCellVal
    Worksheets("Tracker Summaries").Cells(9, summCol).Interior.Color = goodCellColor
    
End Sub


'Search and Grab Serial Number list information
Sub GrabSnInfo(ByVal sEnd As Long, ByVal rData As Range, ByVal summData As Range, ByVal summCol As Integer, ByVal clrBlnk As Long, ByVal clrBlack As Long, ByVal clrRed As Long, ByVal clrOrange As Long, ByVal clrPink As Long, ByVal clrBlue As Long, ByVal clrGreen As Long, ByVal clrbrtGreen As Long, ByVal clrPurple As Long, ByVal clrdrkGreen As Long, ByVal clrlghtGreen1 As Long, ByVal clrlghtGreen2 As Long, ByVal clrYellow As Long, ByVal txtYellow As Long)
    
    'Other Variables
    Dim evnInt As Long
    Dim oddInt As Long
    Dim rowNum As Integer
    Dim colNum As Integer
    Dim rInt As Range
    Dim summInt As Range
    Dim rawsnText As String
    Dim snText As String
    Dim snDate As String
    Dim snColor As Long
    Dim clrRefCell As Long
    
    'Reference Color for light and dark greens and purple
    clrRefCell = Worksheets("Tracker Summaries").Cells(3, summCol).Interior.Color

    'Grab SN Info
    evnInt = 10
    oddInt = 11
    'search for counted cell
    For Each rInt In rData
        rowNum = rInt.Row
        colNum = rInt.Column
        
        'Ignore Yellow Launch Serials
        If (rInt.Row = 43) And (rInt.Interior.Color = clrYellow) Then
            GoTo NextLine
        End If
        
        'set public variables
        Set NEOwrks = ActiveWorkbook.Worksheets("NEO 5322121")
        CellsClearAbove = False
        Set CellCurrent = rInt
        If CellCurrent.Row = 38 Then
           Set CellAbove = CellCurrent.Offset(-5, 0)
        ElseIf CellCurrent.Row = 7 Then
            Set CellAbove = CellCurrent
        Else
            Set CellAbove = CellCurrent.Offset(-1, 0)
        End If
        
        'set cells clear above range
        If CellAbove.Row > 38 Then
            Application.ScreenUpdating = False
            NEOwrks.Activate
            Set CellsAboveRange38 = Range(Cells(CellAbove.Row, CellAbove.Column), Cells(38, colNum))
            Set CellsAboveRange7 = Range(Cells(33, colNum), Cells(7, colNum))
            Worksheets("Tracker Summaries").Activate
            Application.ScreenUpdating = True
        ElseIf CellAbove.Row <= 33 Then
            Application.ScreenUpdating = False
            NEOwrks.Activate
            Set CellsAboveRange38 = NEOwrks.Range(Cells(CellAbove.Row, CellAbove.Column), Cells(7, colNum))
            Set CellsAboveRange7 = NEOwrks.Range(Cells(CellAbove.Row, CellAbove.Column), Cells(7, colNum))
            Worksheets("Tracker Summaries").Activate
            Application.ScreenUpdating = True
        End If
        'set cells clear above boolean
        CellsClearAbove = True
        For Each cllsabvRng In CellsAboveRange38
            If Not cllsabvRng.Interior.Color = clrBlnk Then
                CellsClearAbove = False
                Exit For
            End If
        Next cllsabvRng
        For Each cllsabvRng In CellsAboveRange7
            If Not cllsabvRng.Interior.Color = clrBlnk Then
                CellsClearAbove = False
                Exit For
            End If
        Next cllsabvRng
        
        
        
    'Normal Counting Style
        If (rInt.EntireRow.Hidden = False) And (rInt.EntireColumn.Hidden = False) And Not (clrBlnk = rInt.Interior.Color) And CellsClearAbove And Not (clrBlack = rInt.Interior.Color) Then
            'SN Text, Date, and Color code
            rawsnText = Worksheets("NEO 5322121").Cells(6, colNum).Value
            snText = Right(rawsnText, (Len(rawsnText) - 5))
            snColor = Worksheets("NEO 5322121").Cells(rowNum, colNum).Interior.Color
            snDate = rInt.Value
            'Pink Slow Moving SN Color and Date
            If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrPink Then
                snColor = clrPink
                snDate = Worksheets("NEO 5322121").Cells(52, colNum).Value
                If snDate = "" Then
                    snDate = rInt.Value
                End If
            End If
            'Date N/A
            If snDate = "" Then
                snDate = "---"
            End If
            'Iterate summary column
            For Each summInt In summData
            
            
            'dark greens
                If clrRefCell = clrdrkGreen Then
                    'even numbered cells (SN)
                    If (rInt.Interior.Color = clrdrkGreen) And (summInt.Row = evnInt) Then
                        summInt.Cells(1, 1).Value = snText
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        evnInt = evnInt + 2
                    'odd numbered cells (date)
                    ElseIf (rInt.Interior.Color = clrdrkGreen) And (summInt.Row = oddInt) Then
                        summInt.Cells(1, 1).Value = snDate
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        oddInt = oddInt + 2
                        Exit For
                    End If
                End If
            
            
            'purples
                If clrRefCell = clrPurple Then
                    'even numbered cells (SN)
                    If ((rInt.Interior.Color = clrPurple) Or (rInt.Interior.Color = clrGreen) Or (rInt.Interior.Color = clrOrange)) And (summInt.Row = evnInt) Then
                        summInt.Cells(1, 1).Value = snText
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        evnInt = evnInt + 2
                    'odd numbered cells (date)
                    ElseIf ((rInt.Interior.Color = clrPurple) Or (rInt.Interior.Color = clrGreen) Or (rInt.Interior.Color = clrOrange)) And (summInt.Row = oddInt) Then
                        summInt.Cells(1, 1).Value = snDate
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        oddInt = oddInt + 2
                        Exit For
                    End If
                End If
                
                
            'special (pale) greens
                If clrRefCell = RGB(175, 255, 175) Then
                    'even numbered cells (SN)
                    If ((rInt.Interior.Color = clrdrkGreen) Or (rInt.Interior.Color = clrGreen) Or (rInt.Interior.Color = clrOrange)) And (summInt.Row = evnInt) Then
                        summInt.Cells(1, 1).Value = snText
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        evnInt = evnInt + 2
                    'odd numbered cells (date)
                    ElseIf ((rInt.Interior.Color = clrdrkGreen) Or (rInt.Interior.Color = clrGreen) Or (rInt.Interior.Color = clrOrange)) And (summInt.Row = oddInt) Then
                        summInt.Cells(1, 1).Value = snDate
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        oddInt = oddInt + 2
                        Exit For
                    End If
                End If
            
            
            'normal greens
                If (clrRefCell = clrBlnk) Or (clrRefCell = clrRed) Then
                    'even numbered cells (SN)
                    If Not (rInt.Interior.Color = clrdrkGreen) And Not (rInt.Interior.Color = clrPurple) And (summInt.Row = evnInt) Then
                        summInt.Cells(1, 1).Value = snText
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        evnInt = evnInt + 2
                    'odd numbered cells (date)
                    ElseIf Not (rInt.Interior.Color = clrdrkGreen) And Not (rInt.Interior.Color = clrPurple) And (summInt.Row = oddInt) Then
                        summInt.Cells(1, 1).Value = snDate
                        summInt.Cells(1, 1).Interior.Color = snColor
                        'Yellow Engine Set SN Color
                        If Worksheets("NEO 5322121").Cells(6, colNum).Interior.Color = clrYellow Then
                            summInt.Font.Color = txtYellow
                            summInt.Font.Bold = True
                        End If
                        oddInt = oddInt + 2
                        Exit For
                    End If
                End If
            
            
            Next summInt
        End If
            
            
NextLine:
    Next rInt

End Sub


