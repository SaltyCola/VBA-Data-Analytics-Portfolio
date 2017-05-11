Attribute VB_Name = "TC_Charts"
Public Sub Fix_ChartsWIPSheetReference()

    Dim c As Range
    Dim strFind As String
    Dim pos As Integer
    
    strFind = "[Testing 2-27.xlsm]"
    
    For Each c In ActiveSheet.Range(Cells(1, 1), Cells(475, 54))
        pos = InStr(c.Formula, strFind)
        If pos <> 0 Then
            c.Formula = Left(c.Formula, (pos - 1)) & Right(c.Formula, (Len(c.Formula) - ((pos - 1) + Len(strFind))))
        End If
    Next c

End Sub

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
    lBar.UpdateLoadingBar "Updating Charts 1 of 9", 1, 10
    Call Trend_SlowVReg
    lBar.UpdateLoadingBar "Updating Charts 2 of 9", 2, 10
    Call Trend_TotalBlankOps
    lBar.UpdateLoadingBar "Updating Charts 3 of 9", 3, 10
    Call Trend_NotSeenIn48hrs
    lBar.UpdateLoadingBar "Updating Charts 4 of 9", 4, 10
    Call Table_AverageDaysToDeliver
    lBar.UpdateLoadingBar "Updating Charts 5 of 9", 5, 10
    Call Table_DeliveriesMTD
    lBar.UpdateLoadingBar "Updating Charts 6 of 9", 6, 10
    Call Table_SetTrailingEdge
    lBar.UpdateLoadingBar "Updating Charts 7 of 9", 7, 10
    Call Table_SummarySheetVertical
    lBar.UpdateLoadingBar "Updating Charts 8 of 9", 8, 10
    Call Statistics_PAAPartStatus
    lBar.UpdateLoadingBar "Updating Charts 9 of 9", 9, 10
    Call Statistics_PWAPartStatus
    lBar.UpdateLoadingBar "Updating Charts Complete", 10, 10
    
    'activate charts sheet
    SheetCharts.Activate

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
    iOps = 29 'ignore Launch totals
    sumSlow = 0
    sumReg = 0
    
    'activate charts sheet
    SheetCharts.Activate
    
    'only move entries to new index if the date at maximum days index is not today
    If Cells(58, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(58, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
            Cells(59, (i + 1)).Value = Cells(59, (i + 1 + iDateDelta)).Value
            Cells(60, (i + 1)).Value = Cells(60, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
            If Cells(59, (i + 1)).Value = "" And Cells(60, (i + 1)).Value = "" Then
                Cells(59, (i + 1)).Value = Cells(59, (i)).Value
                Cells(60, (i + 1)).Value = Cells(60, (i)).Value
            End If
        Next i
    End If
    
    'Sum slow and regular movement parts
    For i = 1 To iOps
        sumSlow = sumSlow + Cells(4, (i + 1)).Value
        sumReg = sumReg + Cells(5, (i + 1)).Value
    Next i
    
    'fill in Today's new data
    Cells(58, (iDays + 1)).Value = dToday
    Cells(59, (iDays + 1)).Value = sumSlow
    Cells(60, (iDays + 1)).Value = sumReg

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
    If Cells(113, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(113, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
            Cells(114, (i + 1)).Value = Cells(114, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
            If Cells(114, (i + 1)).Value = "" Then
                Cells(114, (i + 1)).Value = Cells(114, (i)).Value
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
    Cells(113, (iDays + 1)).Value = dToday
    Cells(114, (iDays + 1)).Value = sumBlanks

End Sub

Public Sub Trend_NotSeenIn48hrs()

    Dim dToday As Date 'Today's Date
    Dim iDateDelta As Integer 'Difference between max index date and today's date
    Dim iDays As Integer 'Number of days the trend will show on the chart
    Dim sumUnseenIn As Integer 'Sum of the parts in WIP that haven't been seen in 2 days In House
    Dim sumUnseenOut As Integer 'Sum of the parts in WIP that haven't been seen in 2 days Outsource
    Dim i As Integer 'iterator
    
    'initialize variables
    dToday = Date
    iDateDelta = 0
    iDays = 30
    sumUnseenIn = 0
    sumUnseenOut = 0
    
    'activate charts sheet
    SheetCharts.Activate
    
    'only move entries to new index if the date at maximum days index is not today
    If Cells(167, (iDays + 1)).Value <> dToday Then
        'grab difference of max index date and dToday
        iDateDelta = CInt(dToday - Cells(167, (iDays + 1)).Value)
        'iterate iDays to move each days' data iDateDelta indexes down (Today is at 30)
        For i = 1 To iDays
            'move each entry down by iDateDelta indexes
                'In House
                Cells(168, (i + 1)).Value = Cells(168, (i + 1 + iDateDelta)).Value
                'Outsource
                Cells(169, (i + 1)).Value = Cells(169, (i + 1 + iDateDelta)).Value
            'if entry is empty, then take entry directly left of it
                'In House
                If Cells(168, (i + 1)).Value = "" Then
                    Cells(168, (i + 1)).Value = Cells(168, (i)).Value
                End If
                'Outsource
                If Cells(169, (i + 1)).Value = "" Then
                    Cells(169, (i + 1)).Value = Cells(169, (i)).Value
                End If
        Next i
    End If
    
    'Sum Unseen parts
    For i = 1 To UBound(ArrWIP)
        'found part who's last date seen is <= dToday - 3 (3 to allow for a min of 2 days unseen)
        If CDate(ArrWIP(i).LastDateSeen) <= CDate(dToday - 3) Then
            'Outsource
            If ArrWIP(i).LastOpCompleted = 14 Or ArrWIP(i).LastOpCompleted = 15 Then
                sumUnseenOut = sumUnseenOut + 1
            ElseIf ArrWIP(i).Notes.ValuesList(3) = "SET" And (ArrWIP(i).LastOpCompleted = 24 Or ArrWIP(i).LastOpCompleted = 25) Then
                sumUnseenOut = sumUnseenOut + 1
            'In House
            Else
                sumUnseenIn = sumUnseenIn + 1
            End If
        End If
    Next i
    
    'fill in Today's new data
    Cells(167, (iDays + 1)).Value = dToday
    Cells(168, (iDays + 1)).Value = sumUnseenIn
    Cells(169, (iDays + 1)).Value = sumUnseenOut

End Sub

Public Sub Table_AverageDaysToDeliver()

    Dim colStart As Long 'Holds the column of the beginning of the newest op list UCs
    Dim colEnd As Long 'Holds the column of the end of the newest op list UCs
    Dim cOpDays2Del As Collection 'Collection of Days to Delivery for the current op row, in order from fastest to slowest.
    Dim arrAverages() As Double 'Holds the average number of days between each op and shipping
    Dim bRed As Boolean 'True: First Redline column has been found
    Dim cntTblRw As Integer 'Counter for which table row to paste into
    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    Dim bFastest200 As Boolean 'True: Still looking for fastest 200 ; False: Only looking for 90 day PAA limit parts
    
    'activate sheet
    SheetCharts.Activate
    
    'initialize variables
    Set cOpDays2Del = New Collection
    bRed = False
    
    'resize array
    ReDim arrAverages(1 To ArrWIP(1).NumberOfOps, 1 To 7) As Double
    For i = 1 To UBound(arrAverages)
        arrAverages(i, 1) = 0
        arrAverages(i, 2) = 0
        arrAverages(i, 3) = 0
        arrAverages(i, 4) = 0
        arrAverages(i, 5) = 0
        arrAverages(i, 6) = 0 'parts PAA'ed within last 90 days
        arrAverages(i, 7) = 0 'PAA'ed within 90 days total
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
    
    'iterate through number of operations
    For i = 1 To ArrWIP(1).NumberOfOps
        'reset boolean
        bFastest200 = True
        'iterate reverse through newest op list UCs
        For i1 = colEnd To colStart Step -1
            'set range
            Set c = SheetShipped.Cells(6 + i, i1)
            '200 UC dates for this Oprow read, so exit i1 for loop
            If arrAverages(i, 5) = 200 Then
                bFastest200 = False
            'non empty, date cell
            ElseIf bFastest200 And Not IsEmpty(c) And IsDate(c.Value) Then
                'increment counter in array
                arrAverages(i, 5) = arrAverages(i, 5) + 1
                'test for placement in collection
                If cOpDays2Del.Count = 0 Then
                    cOpDays2Del.Add (CLng(SheetShipped.Cells(7, i1).Value) - CLng(c.Value))
                Else
                    'iterate collection to find placement
                    For i2 = 1 To cOpDays2Del.Count
                        'first delta found larger than current delta, place before.
                        If cOpDays2Del.Item(i2) > (CLng(SheetShipped.Cells(7, i1).Value) - CLng(c.Value)) Then
                            cOpDays2Del.Add (CLng(SheetShipped.Cells(7, i1).Value) - CLng(c.Value)), , i2
                            'end i2 for loop
                            Exit For
                        'no existing entry larger than current, so add to end of collection
                        Else
                            cOpDays2Del.Add (CLng(SheetShipped.Cells(7, i1).Value) - CLng(c.Value))
                            'end i2 for loop
                            Exit For
                        End If
                    Next i2
                End If
            End If
            'Parts PAA'ed within the last 90 Days
            If (CLng(Date) - CLng(SheetShipped.Cells(45, i1).Value)) <= 90 And Not IsEmpty(SheetShipped.Cells(6 + i, i1)) And IsDate(SheetShipped.Cells(6 + i, i1).Value) Then
                'increment PAA'ed Limit Counter
                arrAverages(i, 7) = arrAverages(i, 7) + 1
                'add PAA'ed limit part to totals
                arrAverages(i, 6) = arrAverages(i, 6) + (CLng(SheetShipped.Cells(7, i1).Value) - CLng(SheetShipped.Cells(6 + i, i1).Value))
            End If
        Next i1
        'add sums of op rows to array
        For i1 = 1 To cOpDays2Del.Count
            '100%
            arrAverages(i, 1) = arrAverages(i, 1) + cOpDays2Del.Item(i1)
            '75%
            If i1 <= CInt((cOpDays2Del.Count) * 3 / 4) Then
                arrAverages(i, 2) = arrAverages(i, 2) + cOpDays2Del.Item(i1)
            End If
            '50%
            If i1 <= CInt((cOpDays2Del.Count) * 2 / 4) Then
                arrAverages(i, 3) = arrAverages(i, 3) + cOpDays2Del.Item(i1)
            End If
            '25%
            If i1 <= CInt((cOpDays2Del.Count) * 1 / 4) Then
                arrAverages(i, 4) = arrAverages(i, 4) + cOpDays2Del.Item(i1)
            End If
        Next i1
        'reset collection
        Set cOpDays2Del = New Collection
    Next i
    
    'calculate averages from totals
    For i = 1 To UBound(arrAverages)
        'prevent dividing by zero
        If arrAverages(i, 5) > 0 Then
            arrAverages(i, 1) = arrAverages(i, 1) / arrAverages(i, 5)
            arrAverages(i, 2) = arrAverages(i, 2) / arrAverages(i, 5)
            arrAverages(i, 3) = arrAverages(i, 3) / arrAverages(i, 5)
            arrAverages(i, 4) = arrAverages(i, 4) / arrAverages(i, 5)
        End If
        'PAA'ed Limit Parts
        If arrAverages(i, 7) > 0 Then
            arrAverages(i, 6) = arrAverages(i, 6) / arrAverages(i, 7)
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
                Cells((225 + cntTblRw), 8).Value = arrAverages(i, 1)
                Cells((225 + cntTblRw), 9).Value = arrAverages(i, 2)
                Cells((225 + cntTblRw), 10).Value = arrAverages(i, 3)
                Cells((225 + cntTblRw), 11).Value = arrAverages(i, 4)
                    'parts reported
                    Cells((225 + cntTblRw), 12).Value = arrAverages(i, 5)
                    'sets reported
                    Cells((225 + cntTblRw), 14).Value = arrAverages(i, 5) / 20
                'PAA'ed within the last 90 days averages
                Cells((225 + cntTblRw), 16).Value = arrAverages(i, 6)
                    'parts reported
                    Cells((225 + cntTblRw), 18).Value = arrAverages(i, 7)
                    'sets reported
                    Cells((225 + cntTblRw), 20).Value = arrAverages(i, 7) / 20
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
    Cells(224, 30).Value = cntParts
    Cells(225, 30).Value = cntParts / 20

End Sub

Public Sub Table_SetTrailingEdge()
'Iterate waterfalled ArrWIP array to grab the first 10 sets (25 UCs per set).
'Then grab the LastOpCompleted from the 25th UC in each set, and paste that Op's
' title in the corresponding Table Row.

    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim cntArrTemp As Integer 'counter for resizing arrTemp array
    Dim arrTemp() As TC_UnitColumn 'temporary array for grabbing trailing edge UCs minus CFD/HCF parts
    Dim tOrange As Long 'Temp color variable for this sub specifically
    
    'initialize variables
    tOrange = RGB(255, 192, 0)
    cntArrTemp = 0
    
    'activate sheet
    SheetCharts.Activate
    
    'iterate arrwip to create array of UCs minus CFD/HCF parts (orange complete in FX)
    For i = 1 To UBound(ArrWIP)
        'exit for loop once 500 parts reached in temp array
        If cntArrTemp = 500 Then
            Exit For
        'count everything but orange FX as last completed
        ElseIf Not (ArrWIP(i).LastOpCompleted = 3 And ArrWIP(i).OperationsList(ArrWIP(i).LastOpCompleted).UCColor = tOrange) Then
            'increment counter
            cntArrTemp = cntArrTemp + 1
            'add to temp array
            ReDim Preserve arrTemp(1 To cntArrTemp)
            'add UC to temp array
            Set arrTemp(cntArrTemp) = ArrWIP(i)
        End If
    Next i
    
    'iterate arrwip for the 25th position of each of the first 10 sets.
    For i = 1 To 20
        'paste 25th entry of each of the first 10 groups of 25
        Cells(227 + i, 28).Value = arrTemp(i * 25).OperationsList(arrTemp(i * 25).LastOpCompleted).Title
        'iterate Average Days to Delivery table for expected ship date
        For i1 = 1 To 29
            'found correct op row in Average Days to Delivery table
            If Cells((i1 + 225), 3).Value = Cells(227 + i, 28).Value Then
                'paste expected delivery date
                Cells(227 + i, 32).Value = Date + CInt(Cells((i1 + 225), 16).Value)
                'exit for loop
                Exit For
            End If
        Next i1
    Next i

End Sub

Public Sub Table_SummarySheetVertical()
'Iterate Slow Vs Regular Table at top of Charts tab and copy into vertical table.

    Dim i As Integer 'iterator
    
    'activate charts sheet
    SheetCharts.Activate
    
    'copy values
    For i = 1 To 29 '30 total, but ignoring "Shipped to NELC" op in SlowVReg Table for now
        Cells((i + 223), 47).Value = Cells(5, (i + 2)).Value
        Cells((i + 223), 50).Value = Cells(4, (i + 2)).Value
    Next i

End Sub

Public Sub Statistics_PAAPartStatus()
' 1. Iterate WIP, QC, and Shipped tabs
' 2. Grab (if applicable) SN, Tab Name, PAA Date, and Year's Week
' 3. Count PAA dates by week of current year

    Dim arrWeeksCount() As Long 'Array to hold the number of UC's per week
    Dim cntWeeks As Integer 'counter for the number of weeks in this current year to size array
    Dim rwPAA As Integer 'current row where PAA dates are located
    Dim bRedline As Boolean 'True: the next redline found will end the iteration
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim d As Date 'iterator
    Dim bD As Boolean 'boolean for d iterator
    Dim strD As String 'string for d iterator
    Dim c As Range 'iterator
    Dim c1 As Range 'iterator
    
    'resize weeks count array
    cntWeeks = 53 '53 weeks to grab 1 whole running year
    ReDim arrWeeksCount(1 To cntWeeks, 1 To 5) As Long 'arr(1: week number, 2: (1=WIP,2=Shipped,3=QC,4=StartWeek,5=EndWeek))
    
    'initialize array
    For i = 1 To cntWeeks
        arrWeeksCount(i, 1) = 0
        arrWeeksCount(i, 2) = 0
        arrWeeksCount(i, 3) = 0
        arrWeeksCount(i, 4) = CLng(Date) - (370 - ((i - 1) * 7))
        arrWeeksCount(i, 5) = CLng(Date) - (370 - ((i - 1) * 7)) + 6
    Next i
    
    'iterate 3 times for three sheets
    For i = 1 To 3
        'activate sheets and initialize boolean
        If i = 1 Then
            SheetWIP.Activate
            bRedline = True
        ElseIf i = 2 Then
            SheetShipped.Activate
            bRedline = False
        ElseIf i = 3 Then
            SheetQC.Activate
            bRedline = True
        End If
        'iterate sn row
        For Each c In Range("6:6")
            'nonfinal redline found
            If Not bRedline And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                bRedline = True
            'final redline found
            ElseIf bRedline And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                Exit For
            'find rwPAA
            ElseIf c.Value = "Process Bucket" Then
                For Each c1 In Range(Cells(1, c.Column), Cells(50, c.Column))
                    If Left(c1.Value, 3) = "PAA" Then
                        rwPAA = c1.Row
                        Exit For
                    End If
                Next c1
            'valid Unit Column
            ElseIf rwPAA <> 0 And Not c.EntireColumn.Hidden And Not IsEmpty(c) Then
                'make sure UC has a PAA date to use
                If IsDate(c.Offset((rwPAA - 6), 0).Value) Then
                    'PAA date is within the running year
                    If c.Offset((rwPAA - 6), 0).Value >= CDate(CLng(Date) - 370) Then '370 gives 53 weeks of 7 days including today
                        'find placement in array
                        For i1 = 1 To cntWeeks
                            'placement found
                            If c.Offset((rwPAA - 6), 0).Value > arrWeeksCount(i1, 4) And c.Offset((rwPAA - 6), 0).Value < arrWeeksCount(i1, 5) Then
                                'increment value to array
                                arrWeeksCount(i1, i) = arrWeeksCount(i1, i) + 1
                                Exit For
                            End If
                        Next i1
                    End If
                End If
            End If
        Next c
    Next i
    
    'activate charts sheet
    SheetCharts.Activate
    
    'write array to chart table
    For i = 1 To UBound(arrWeeksCount)
        'set week string for chart axis
            'reset strD
            strD = ""
            'start date
            strD = Left(CStr(CDate(arrWeeksCount(i, 4))), (Len(CStr(CDate(arrWeeksCount(i, 4)))) - 5)) & " - "
            'end date
            strD = strD & CStr(CDate(arrWeeksCount(i, 5)))
        'fill in values
        Cells(259, (i + 1)).Value = strD
        Cells(260, (i + 1)).Value = arrWeeksCount(i, 1)
        Cells(261, (i + 1)).Value = arrWeeksCount(i, 2)
        Cells(262, (i + 1)).Value = arrWeeksCount(i, 3)
    Next i

End Sub

Public Sub Statistics_PWAPartStatus()
' 1. Iterate WIP, QC, and Shipped tabs
' 2. Grab (if applicable) SN, Tab Name, OX to PWA Date, and Year's Week
' 3. Count PWA dates by week of current year

    Dim arrWeeksCount() As Long 'Array to hold the number of UC's per week
    Dim cntWeeks As Integer 'counter for the number of weeks in this current year to size array
    Dim rwPWA As Integer 'current row where PWA dates are located
    Dim bRedline As Boolean 'True: the next redline found will end the iteration
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim d As Date 'iterator
    Dim bD As Boolean 'boolean for d iterator
    Dim strD As String 'string for d iterator
    Dim c As Range 'iterator
    Dim c1 As Range 'iterator
    
    'resize weeks count array
    cntWeeks = 53 '53 weeks to grab 1 whole running year
    ReDim arrWeeksCount(1 To cntWeeks, 1 To 5) As Long 'arr(1: week number, 2: (1=WIP,2=Shipped,3=QC,4=StartWeek,5=EndWeek))
    
    'initialize array
    For i = 1 To cntWeeks
        arrWeeksCount(i, 1) = 0
        arrWeeksCount(i, 2) = 0
        arrWeeksCount(i, 3) = 0
        arrWeeksCount(i, 4) = CLng(Date) - (370 - ((i - 1) * 7))
        arrWeeksCount(i, 5) = CLng(Date) - (370 - ((i - 1) * 7)) + 6
    Next i
    
    'iterate 3 times for three sheets
    For i = 1 To 3
        'activate sheets and initialize boolean
        If i = 1 Then
            SheetWIP.Activate
            bRedline = True
        ElseIf i = 2 Then
            SheetShipped.Activate
            bRedline = False
        ElseIf i = 3 Then
            SheetQC.Activate
            bRedline = True
        End If
        'iterate sn row
        For Each c In Range("6:6")
            'nonfinal redline found
            If Not bRedline And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                bRedline = True
            'final redline found
            ElseIf bRedline And c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                Exit For
            'find rwPWA
            ElseIf c.Value = "Process Bucket" Then
                For Each c1 In Range(Cells(1, c.Column), Cells(50, c.Column))
                    If Left(c1.Value, 13) = "OX Automation" Or Left(c1.Value, 23) = "OX / Transit Automation" Or Left(c1.Value, 22) = "Outgoing to Automation" Then
                        rwPWA = c1.Row
                        Exit For
                    End If
                Next c1
            'valid Unit Column
            ElseIf rwPWA <> 0 And Not c.EntireColumn.Hidden And Not IsEmpty(c) Then
                'make sure UC has a PWA date to use
                If IsDate(c.Offset((rwPWA - 6), 0).Value) Then
                    'PWA date is within the running year
                    If c.Offset((rwPWA - 6), 0).Value >= CDate(CLng(Date) - 370) Then '370 gives 53 weeks of 7 days including today
                        'find placement in array
                        For i1 = 1 To cntWeeks
                            'placement found
                            If c.Offset((rwPWA - 6), 0).Value > arrWeeksCount(i1, 4) And c.Offset((rwPWA - 6), 0).Value < arrWeeksCount(i1, 5) Then
                                'increment value to array
                                arrWeeksCount(i1, i) = arrWeeksCount(i1, i) + 1
                                Exit For
                            End If
                        Next i1
                    End If
                End If
            End If
        Next c
    Next i
    
    'activate charts sheet
    SheetCharts.Activate
    
    'write array to chart table
    For i = 1 To UBound(arrWeeksCount)
        'set week string for chart axis
            'reset strD
            strD = ""
            'start date
            strD = Left(CStr(CDate(arrWeeksCount(i, 4))), (Len(CStr(CDate(arrWeeksCount(i, 4)))) - 5)) & " - "
            'end date
            strD = strD & CStr(CDate(arrWeeksCount(i, 5)))
        'fill in values
        Cells(368, (i + 1)).Value = strD
        Cells(369, (i + 1)).Value = arrWeeksCount(i, 1)
        Cells(370, (i + 1)).Value = arrWeeksCount(i, 2)
        Cells(371, (i + 1)).Value = arrWeeksCount(i, 3)
    Next i

End Sub

Public Sub ExportChartsToPowerPoint()

'    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B6:AE55"), True, False) 'slide 1
'    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AG6:AI11"), False, True, 1)
'    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B61:AE110"), False, False) 'slide 2
    
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("V222:AK248"), True, False, 1) 'slide 1 (shift down slightly)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AJ222:BC253"), False, False) 'slide 2
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B222:G256"), False, False, 3) 'slide 3 (1/2)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("P222:V256"), False, True, 3) 'slide 3 (2/2)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B317:AE366"), False, False) 'slide 4
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AG317:AP324"), False, True, 4)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B426:AE475"), False, False) 'slide 5
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AG426:AP433"), False, True, 5)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AR317:BA324"), False, False, 6) 'slide 6
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("AR426:BA433"), False, True, 6)
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B115:AE164"), False, False) 'slide 7
    Call ExcelRangeToPowerPoint(ThisWorkbook.ActiveSheet.Range("B170:AE219"), False, False) 'slide 8

End Sub

Public Sub ExcelRangeToPowerPoint(ByRef rng As Range, ByVal boolNewPresentation As Boolean, ByVal boolSupplementaryTable As Boolean, Optional ByVal intSlideNumber As Integer)

    Dim PowerPointApp As Object
    Dim myPresentation As Object
    Dim mySlide As Object
    Dim myShape As Object
    Dim wrksCharts As Worksheet
    
    'get worksheet
    Set wrksCharts = ActiveSheet
    
    'select cell C1
    Cells(1, 3).Select
    
    'create an instance of PowerPoint
    On Error Resume Next
        
        'get powerpoint reference
        Set PowerPointApp = GetObject(Class:="PowerPoint.Application")
        
        'clear the error between errors
        Err.Clear
        
        'if PowerPoint is not already open then open PowerPoint
        If PowerPointApp Is Nothing Then
            Set PowerPointApp = CreateObject(Class:="PowerPoint.Application")
        End If
        
        'handle if the PowerPoint Application is not found
        If Err.Number = 429 Then
            MsgBox "PowerPoint could not be found, aborting."
            Exit Sub
        End If
    
    On Error GoTo 0
    
    'optimize code
    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False
    
    'create new presentation
    If boolNewPresentation Then
        Set myPresentation = PowerPointApp.Presentations.Add
    'presentation already there
    Else
        Set myPresentation = PowerPointApp.ActivePresentation
    End If
    
    'add a slide to the presentation
    If Not boolSupplementaryTable Then
        Set mySlide = myPresentation.Slides.Add((myPresentation.Slides.Count + 1), 12) '12 = ppLayoutBlank
    'supplementary tables
    ElseIf boolSupplementaryTable Then
        Set mySlide = myPresentation.Slides(intSlideNumber)
    End If
    
    'copy Excel range
    wrksCharts.Activate
    rng.Copy
    
    'paste to PowerPoint and position
    myPresentation.Application.Activate
    mySlide.Shapes.PasteSpecial DataType:=2 '2 = ppPasteEnhancedMetafile
    Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
        
        'set size and position
            'regular charts and tables
            If Not boolSupplementaryTable And (intSlideNumber <> 3) And (intSlideNumber <> 6) Then
                'wider than tall
                If myShape.Width > myShape.Height Then
                    myShape.Height = myShape.Height * 720 / myShape.Width
                    myShape.Width = 720
                    If myShape.Width < 720 Then
                        myShape.Left = 360 - (myShape.Width / 2)
                    Else
                        myShape.Left = 0
                    End If
                    If myShape.Height < 540 Then
                        myShape.Top = 270 - (myShape.Height / 2)
                    Else
                        myShape.Top = 0
                    End If
                'taller than wide
                Else
                    myShape.Height = 540
                    myShape.Width = myShape.Width * 540 / myShape.Height
                    If myShape.Width < 720 Then
                        myShape.Left = 360 - (myShape.Width / 2)
                    Else
                        myShape.Left = 0
                    End If
                    If myShape.Height < 540 Then
                        myShape.Top = 270 - (myShape.Height / 2)
                    Else
                        myShape.Top = 0
                    End If
                End If
            
            'supplementary tables
            ElseIf boolSupplementaryTable And (intSlideNumber <> 3) And (intSlideNumber <> 6) Then
                'reduce size
                myShape.Height = myShape.Height * 0.59
                myShape.Width = myShape.Width * 0.59
                'top right corner of slide
                myShape.Top = 1
                myShape.Left = 720 - myShape.Width
            
            'First half of "Average Days to Deliver" Table
            ElseIf Not boolSupplementaryTable And (intSlideNumber = 3) Then
                'taller than wide
                    myShape.Height = 540
                    myShape.Width = myShape.Width * 540 / myShape.Height
                    myShape.Left = 360 - myShape.Width
                    myShape.Top = 270 - (myShape.Height / 2)
            
            'Second half of "Average Days to Deliver" Table
            ElseIf boolSupplementaryTable And (intSlideNumber = 3) Then
                'taller than wide
                    myShape.Height = 540
                    myShape.Width = myShape.Width * 540 / myShape.Height
                    myShape.Left = 360 - 2 '(-2) for better alignment
                    myShape.Top = 270 - (myShape.Height / 2)
            
            '"PAA Breakdown Totals Year to Date" Table
            ElseIf Not boolSupplementaryTable And (intSlideNumber = 6) Then
                'taller than wide
                    myShape.Height = (540 / 3)
                    myShape.Width = myShape.Width * (540 / 3) / myShape.Height
                    myShape.Left = 360 - (myShape.Width / 2)
                    myShape.Top = ((540 / 4) * 1) - (myShape.Height / 2)
            
            '"PWA Breakdown Totals Year to Date" Table
            ElseIf boolSupplementaryTable And (intSlideNumber = 6) Then
                'taller than wide
                    myShape.Height = (540 / 3)
                    myShape.Width = myShape.Width * (540 / 3) / myShape.Height
                    myShape.Left = 360 - (myShape.Width / 2)
                    myShape.Top = ((540 / 4) * 3) - (myShape.Height / 2)
            
            End If
    
    'shift slide 1's shape down slightly
    If intSlideNumber = 1 Then
        myShape.Top = myShape.Top + 8
    End If
    
    'make PowerPoint visible and active
    PowerPointApp.Visible = True
    PowerPointApp.Activate
    
    'clear clipboard
    Application.CutCopyMode = False
    
    'restart screen updating
    Application.ScreenUpdating = True
    ActiveWindow.DisplayGridlines = True

End Sub
