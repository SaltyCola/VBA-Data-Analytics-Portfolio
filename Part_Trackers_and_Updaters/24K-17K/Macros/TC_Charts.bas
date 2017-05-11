Attribute VB_Name = "TC_Charts"
Public Sub GenerateCharts()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
    bReadWIP = True
    bInitializeEventHandlerCollections = True
    bInitializeUserForms = True
    bWaterfallSort = True
    bCompleteSummary = True
    
    'call T.U.P.
    Call TUP_Initialize
    
    'call chart subs (overarching loading bar)
    lBar.UpdateLoadingBar "Updating Charts 1 of 7", 1, 8
    lBar.UpdateLoadingBar "Updating Charts 2 of 7", 2, 8
    Call Trend_SlowVReg
    lBar.UpdateLoadingBar "Updating Charts 3 of 7", 3, 8
    Call Trend_TotalBlankOps
    lBar.UpdateLoadingBar "Updating Charts 4 of 7", 4, 8
    Call Trend_NotSeenIn48hrs
    lBar.UpdateLoadingBar "Updating Charts 5 of 7", 5, 8
    Call Table_AverageDaysToDeliver
    lBar.UpdateLoadingBar "Updating Charts 6 of 7", 6, 8
    Call Table_DeliveriesMTD
    lBar.UpdateLoadingBar "Updating Charts 7 of 7", 7, 8
    Call Table_SetTrailingEdge
    lBar.UpdateLoadingBar "Updating Charts Complete", 8, 8

End Sub

Public Sub Trend_SlowVReg()

    Dim dToday As Date 'Today's Date
    Dim iDateDelta As Integer 'Difference between max index date and today's date
    Dim iDays As Integer 'Number of days the trend will show on the chart
    Dim iOps As Integer 'Number of ops in slow vs. regular chart
    Dim sumSlow As Integer 'Sum of the slow parts in WIP
    Dim sumReg As Integer 'Sum of the regular movement parts in WIP
    Dim i As Integer 'iterator
    
    'initialize variables
    dToday = Date
    iDateDelta = 0
    iDays = 30
    iOps = 30
    sumSlow = 0
    sumReg = 0
    
    'activate charts sheet
    SheetCharts.Activate
    
    'only move entries to new index if the date at maximum days index is not today
    If Cells(45, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(45, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
            Cells(46, (i + 1)).Value = Cells(46, (i + 1 + iDateDelta)).Value
            Cells(47, (i + 1)).Value = Cells(47, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
            If Cells(46, (i + 1)).Value = "" And Cells(47, (i + 1)).Value = "" Then
                Cells(46, (i + 1)).Value = Cells(46, (i)).Value
                Cells(47, (i + 1)).Value = Cells(47, (i)).Value
            End If
        Next i
    End If
    
    'Sum slow and regular movement parts
    For i = 1 To iOps
        sumSlow = sumSlow + Cells(4, (i + 1)).Value
        sumReg = sumReg + Cells(5, (i + 1)).Value
    Next i
    
    'fill in Today's new data
    Cells(45, (iDays + 1)).Value = dToday
    Cells(46, (iDays + 1)).Value = sumSlow
    Cells(47, (iDays + 1)).Value = sumReg

End Sub
         
Public Sub Trend_TotalBlankOps()

    Dim dToday As Date 'Today's Date
    Dim iDateDelta As Integer 'Difference between max index date and today's date
    Dim iDays As Integer 'Number of days the trend will show on the chart
    Dim sumBlanks As Integer 'Sum of the blank ops in WIP
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    
    'initialize variables
    dToday = Date
    iDateDelta = 0
    iDays = 30
    sumBlanks = 0
    
    'activate charts sheet
    SheetCharts.Activate
    
    'only move entries to new index if the date at maximum days index is not today
    If Cells(87, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(87, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
            Cells(88, (i + 1)).Value = Cells(88, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
            If Cells(88, (i + 1)).Value = "" Then
                Cells(88, (i + 1)).Value = Cells(88, (i)).Value
            End If
        Next i
    End If
    
    'Sum Blanks
    For i1 = 1 To UBound(ArrWIP)
        'iterate from "Launch" to Last Op Completed
        For i2 = ArrWIP(i1).NumberOfOps To ArrWIP(i1).LastOpCompleted Step -1
            'op is not hidden
            If ArrWIP(i1).OperationsList(i2).Enabled Then
                'blank found (zero-date = blank)
                If CDate(ArrWIP(i1).OperationsList(i2).UCDate) = CDate(0) Then
                    sumBlanks = sumBlanks + 1
                End If
            End If
        Next i2
    Next i1
    
    'fill in Today's new data
    Cells(87, (iDays + 1)).Value = dToday
    Cells(88, (iDays + 1)).Value = sumBlanks

End Sub

Public Sub Trend_NotSeenIn48hrs()

    Dim dToday As Date 'Today's Date
    Dim iDateDelta As Integer 'Difference between max index date and today's date
    Dim iDays As Integer 'Number of days the trend will show on the chart
    Dim sumUnseen As Integer 'Sum of the parts in WIP that haven't been seen in 2 days
    Dim i As Integer 'iterator
    
    'initialize variables
    dToday = Date
    iDateDelta = 0
    iDays = 30
    sumUnseen = 0
    
    'activate charts sheet
    SheetCharts.Activate
    
    'only move entries to new index if the date at maximum days index is not today
    If Cells(128, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(128, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
            Cells(129, (i + 1)).Value = Cells(129, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
            If Cells(129, (i + 1)).Value = "" Then
                Cells(129, (i + 1)).Value = Cells(129, (i)).Value
            End If
        Next i
    End If
    
    'Sum Unseen parts
    For i = 1 To UBound(ArrWIP)
        'found part who's last date seen is <= dToday - 3 (3 to allow for a min of 2 days unseen)
        If CDate(ArrWIP(i).LastDateSeen) <= CDate(dToday - 3) Then
            sumUnseen = sumUnseen + 1
        End If
    Next i
    
    'fill in Today's new data
    Cells(128, (iDays + 1)).Value = dToday
    Cells(129, (iDays + 1)).Value = sumUnseen

End Sub

Public Sub Table_AverageDaysToDeliver()

    Dim colStart As Long 'Holds the column of the beginning of the newest op list UCs
    Dim colEnd As Long 'Holds the column of the end of the newest op list UCs
    Dim arrAverages() As Double 'Holds the average number of days between each op and shipping
    Dim bRed As Boolean 'True: Redline column has been found
    Dim cntTblRw As Integer 'Counter for which table row to paste into
    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    
    'activate sheet
    SheetCharts.Activate
    
    'initialize variables
    bRed = False
    
    'resize array
    ReDim arrAverages(1 To ArrWIP(1).NumberOfOps, 1 To 2) As Double
    For i = 1 To UBound(arrAverages)
        arrAverages(i, 1) = 0
        arrAverages(i, 2) = 0
    Next i
    
    'iterate Shipped tab
    For Each c In SheetShipped.Range("6:6")
        'found redline (where newest op list UCs begin) and count newest UCs)
        If Not bRed And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            colStart = c.Column + 1
            bRed = True
        'second redline ends the loop
        ElseIf bRed And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            colEnd = c.Column - 1
            Exit For
        End If
    Next c
    
    'iterate reverse through newest op list UCs
    For i = colEnd To colStart Step -1
        'set range
        Set c = SheetShipped.Cells(6, i)
        'non empty SN cell
        If Not IsEmpty(c) Then
            'add values to array
            For i1 = 1 To UBound(arrAverages)
                'ignore empty cells
                If Not IsEmpty(c.Offset(i1, 0)) Then
                    'only grab most recent 200 at maximum
                    If arrAverages(i1, 2) < 200 Then
                        'sum of days to delivered per op row as compared to op row 1
                        arrAverages(i1, 1) = arrAverages(i1, 1) + (CLng(c.Offset(1, 0).Value) - CLng(c.Offset(i1, 0).Value))
                        'total
                        arrAverages(i1, 2) = arrAverages(i1, 2) + 1
                    End If
                End If
            Next i1
        End If
    Next i
    
    'calculate averages from totals
    For i = 1 To UBound(arrAverages)
        'prevent dividing by zero
        If arrAverages(i, 2) > 0 Then
            arrAverages(i, 1) = arrAverages(i, 1) / arrAverages(i, 2)
        End If
    Next i
    
    'initialize counter
    cntTblRw = 0
    
    'iterate averages array for pasting to table
    For i = 1 To UBound(arrAverages)
        'row is enabled and not the first row
        If ArrWIP(1).OperationsList(i).Enabled And i <> 1 Then
            'increment counter
            cntTblRw = cntTblRw + 1
            'paste to table
                'averages
                Cells((169 + cntTblRw), 10).Value = arrAverages(i, 1)
                'parts reported
                Cells((169 + cntTblRw), 12).Value = arrAverages(i, 2)
                'sets reported
                Cells((169 + cntTblRw), 14).Value = arrAverages(i, 2) / 20
        End If
    Next i

End Sub

Public Sub Table_DeliveriesMTD()

    Dim dMonth As Integer 'today's month
    Dim dYear As Integer 'today's Year
    Dim cntParts As Long 'Number of Parts delivered with this month/year combo
    Dim bRed As Boolean 'True: Redline column has been found
    Dim c As Range 'iterator
    
    'activate sheet
    SheetShipped.Activate
    
    'initialize variables
    dMonth = Month(Date)
    dYear = Year(Date)
    cntParts = 0
    bRed = False
    
    'iterate Shipped tab's delivery row (row 7)
    For Each c In Range("7:7")
        'value is a date
        If IsDate(c.Value) And c.Interior.Color = RGB(146, 208, 80) Then
            'parts delivered this month
            If Month(CDate(c.Value)) = dMonth And Year(CDate(c.Value)) = dYear And bRed Then
                'increment count
                cntParts = cntParts + 1
            End If
        'first redline found, change boolean
        ElseIf c.EntireColumn.Interior.Color = RGB(255, 0, 0) And Not bRed Then
            bRed = True
        'end loop at second red line
        ElseIf c.EntireColumn.Interior.Color = RGB(255, 0, 0) And bRed Then
            Exit For
        End If
    Next c
    
    'activate sheet
    SheetCharts.Activate
    
    'paste to table
    Cells(170, 27).Value = cntParts
    Cells(171, 27).Value = cntParts / 20

End Sub

Public Sub Table_SetTrailingEdge()
'Iterate waterfalled ArrWIP array to grab the first 10 sets (25 UCs per set).
'Then grab the LastOpCompleted from the 25th UC in each set, and paste that Op's
' title in the corresponding Table Row.

    Dim i As Integer 'iterator
    
    'activate sheet
    SheetCharts.Activate
    
    'iterate arrwip for the 25th position of each of the first 10 sets.
    For i = 1 To 10
        'paste 25th entry of each of the first 10 groups of 25
        Cells(172 + i, 25).Value = ArrWIP(i * 25).OperationsList(ArrWIP(i * 25).LastOpCompleted).Title
    Next i

End Sub
