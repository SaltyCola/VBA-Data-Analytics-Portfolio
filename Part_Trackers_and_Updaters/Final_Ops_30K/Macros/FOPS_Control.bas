Attribute VB_Name = "FOPS_Control"
Public ArrSNRows() As FOPS_SNRow 'Array of all SN Row objects
Public SheetFOPS As Worksheet 'worksheet variable for reading
Public SheetFOPS_BACKUP As Worksheet 'backup copy of worksheet to prevent total loss of information
Public SheetSummary As Worksheet 'worksheet variable for summary page
Public SNRowsTotal As Integer 'total number of SN Rows
Public FirstListRow As Integer 'first row where SN Row list starts
Public LastListRow As Integer 'last row where SN Row list ends
Public ColumnComments As Integer 'final column in the table
Public ColumnOpsStart As Integer 'column where operations begin
Public ColumnOpsEnd As Integer 'column where operations end
Public lBar As FOPS_LoadingBar 'ainmated progress bar
'subroutine booleans
    Public bReadSNRows As Boolean 'directs subdirector on whether to run or not
    Public bWaterfallSNRows As Boolean 'directs subdirector on whether to run or not
    Public bCreateBackupSheetCopy As Boolean 'directs subdirector on whether to run or not
    Public bClearSNRows As Boolean 'directs subdirector on whether to run or not
    Public bWriteSummary As Boolean 'directs subdirector on whether to run or not
    Public bWriteSNRows As Boolean 'directs subdirector on whether to run or not
    Public bFixDaysAtPWA As Boolean 'directs subdirector on whether to run or not
    Public bDeleteBackupSheetCopy As Boolean 'directs subdirector on whether to run or not

Public Sub SummaryFOPS()

    'initialize booleans
    bReadSNRows = False
    bWaterfallSNRows = False
    bCreateBackupSheetCopy = False
    bClearSNRows = False
    bWriteSNRows = False
    bFixDaysAtPWA = False
    bDeleteBackupSheetCopy = False
    
    'set booleans
    bReadSNRows = True
    bWriteSummary = True
    
    'call sub director
    Call SubDirector

End Sub

Public Sub WaterfallSNList()

    'initialize booleans
    bReadSNRows = False
    bWaterfallSNRows = False
    bCreateBackupSheetCopy = False
    bClearSNRows = False
    bWriteSNRows = False
    bFixDaysAtPWA = False
    bDeleteBackupSheetCopy = False
    
    'set booleans
    bReadSNRows = True
    bWaterfallSNRows = True
    bCreateBackupSheetCopy = True
    bClearSNRows = True
    bWriteSNRows = True
    bFixDaysAtPWA = True
    bDeleteBackupSheetCopy = True
    
    'call sub director
    Call SubDirector

End Sub

Public Sub SubDirector()

    Dim c As Range 'iterator
    
    'get worksheet
    Set SheetFOPS = ActiveWorkbook.Worksheets("Final Operations")
    Set SheetSummary = ActiveWorkbook.Worksheets("Summary")
    
    'initialize loading bar
    Set lBar = New FOPS_LoadingBar
    
    'screen updating off
    Application.ScreenUpdating = False
    
    'call subs
    If bReadSNRows Then: Call ReadSNRows
    
    If bWaterfallSNRows Then: Call WaterfallSNRows
    
    'create backup FOps sheet copy
    If bCreateBackupSheetCopy Then: Call CreateBackupSheetCopy
    
    If bClearSNRows Then: Call ClearSNRows
    
    If bWriteSummary Then: Call WriteSummary
    
    If bWriteSNRows Then: Call WriteSNRows
    
    'fix Days at PWA column Formulas
    If bFixDaysAtPWA Then: Call FixDaysAtPWA
    
    'delete backup FOps sheet copy
    If bDeleteBackupSheetCopy Then: Call DeleteBackupSheetCopy
    
    'screen updating on
    Application.ScreenUpdating = True

End Sub

Public Sub ReadSNRows()

    Dim c As Range 'iterator
    Dim iOps As Integer 'counter for operation columns
    Dim tSNRow As FOPS_SNRow 'temp SNRow object
    
    'activate sheet
    SheetFOPS.Activate
    
    'find first list row
    For Each c In SheetFOPS.Range("A:A")
        If c.Value = "Program" Then
            FirstListRow = c.Row + 1
            Exit For
        End If
    Next c
    
    'find last list row
    For Each c In SheetFOPS.Range("A:A")
        If c.Row > FirstListRow And IsEmpty(c) Then
            LastListRow = c.Row - 1
            Exit For
        End If
    Next c
    
    'find number of operations columns
    ColumnOpsStart = 0
    ColumnOpsEnd = 0
    For Each c In SheetFOPS.Range("2:2")
        If c.Value = "Days at PWA" Then
            ColumnOpsStart = c.Column + 1
        ElseIf c.Interior.Color = RGB(0, 0, 0) Then
            ColumnOpsEnd = c.Column - 1
            Exit For 'only looks for first black column
        End If
    Next c
    
    'find final column ("comments")
    For Each c In SheetFOPS.Range("2:2")
        If c.Value = "Comments" Then: ColumnComments = c.Column
    Next c
    
    'add SN Row objects to array
    SNRowsTotal = 0
    For Each c In SheetFOPS.Range(Cells(FirstListRow, 1), Cells(LastListRow, 1))
        'initialize temp SNRow object
        Set tSNRow = New FOPS_SNRow
        'increment counter
        SNRowsTotal = SNRowsTotal + 1
        'resize array
        ReDim Preserve ArrSNRows(1 To SNRowsTotal) As FOPS_SNRow
        'get data
        tSNRow.DateFromPWA = c.Offset(0, (ColumnOpsStart - 3)).Value '(-1 because offset, - 2 because column before columnopsstart)
        tSNRow.NumberOfOperations = (1 + ColumnOpsEnd) - ColumnOpsStart
        tSNRow.ColOpsStart = ColumnOpsStart
        tSNRow.ColOpsEnd = ColumnOpsEnd
        tSNRow.GrabData c, ColumnComments 'start in c.column and go to Comments column to the right
        'add to array
        Set ArrSNRows(SNRowsTotal) = tSNRow
        
        'update loading bar
        lBar.UpdateLoadingBar "Reading SN Rows from table...", (c.Row - (FirstListRow - 1)), (LastListRow - (FirstListRow - 1))
    
    Next c

End Sub

Public Sub WaterfallSNRows()

    Dim tArrSNRows() As FOPS_SNRow 'temp array to waterfall then give back to ArrSNRows
    Dim tarrOpGroup() As Variant 'temp array to waterfall each waterfallIndex group by PWA received Date
    Dim cntOpGroup As Integer 'counter for adding SNRows to temp group array from original array
    Dim cntSNRow As Integer 'counter for adding SNRows to temp SNRow array from op group arrays
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    Dim i3 As Integer 'iterator
    Dim maxInd As Integer 'maximum waterfall index to waterfall on top
    
    'activate sheet
    SheetFOPS.Activate
    
    'initialize size of temp array
    ReDim tArrSNRows(1 To UBound(ArrSNRows)) As FOPS_SNRow
    
    'Initialize counters
    maxInd = 0
    cntSNRow = 0
    
    'grab max waterfall index
    maxInd = ArrSNRows(1).NumberOfOperations + 1
    
    'waterfall SNRows into temp array (iterate in reverse)
    For i1 = maxInd To 0 Step -1
        
        'Initialize counter
        cntOpGroup = 0
        
        'reset tarrOpGroup
        ReDim tarrOpGroup(1 To 1) As Variant
        
        'look for entries with this index
        For i2 = 1 To UBound(ArrSNRows)
            'entry found and unsorted
            If ArrSNRows(i2).WaterfallIndex = i1 And Not ArrSNRows(i2).WSorted Then
                'increment counter for adding to temp array
                cntOpGroup = cntOpGroup + 1
                'resize group array
                ReDim Preserve tarrOpGroup(1 To cntOpGroup) As Variant
                'add to group array
                Set tarrOpGroup(cntOpGroup) = ArrSNRows(i2)
                'change sorted variable in original
                ArrSNRows(i2).WSorted = True
            End If
        Next i2
        
        'only continue with following loops if cntOpGroup <> 0
        If cntOpGroup <> 0 Then
        
            'Reset group array's wSorted booleans
            For i2 = 1 To UBound(tarrOpGroup)
                tarrOpGroup(i2).WSorted = False
            Next i2
            
            'Sort based on PWA Received Date
            For i2 = 2 To UBound(tarrOpGroup) 'index to move (always skip number 1)
                'iterate group array again to find placement for i2 index
                 For i3 = 1 To UBound(tarrOpGroup) 'index for placement
                    'Do nothing and move on to next i2 if i3-loop reaches i3 = i2
                    If i3 = i2 Then
                        'end i3 for loop
                        Exit For
                    'i3 does not = i2, so compare dates (larger date is the later date)
                    ElseIf i3 <> i2 Then
                        If tarrOpGroup(i3).DateFromPWA > tarrOpGroup(i2).DateFromPWA Then
                            'move i2 to i3's position
                            tarrOpGroup = MoveArrayEntry(tarrOpGroup, i2, i3)
                            'end i3 for loop
                            Exit For
                        End If
                    End If
                 Next i3
            Next i2
            
            'add op group array entries to temp SN row array
            For i2 = 1 To UBound(tarrOpGroup)
                'increment counter
                cntSNRow = cntSNRow + 1
                'add op group entry to tArrSNRows
                Set tArrSNRows(cntSNRow) = tarrOpGroup(i2)
            Next i2
            
        End If
        
        'update loading bar
        lBar.UpdateLoadingBar "Waterfalling SN Rows...", ((maxInd + 1) - i), (maxInd + 1)
    
    Next i1
    
    'set original array to temp array
    ArrSNRows() = tArrSNRows()
    
    'iterate original array to set ListIndex and alter V-lookups
    For i1 = 1 To UBound(ArrSNRows)
        ArrSNRows(i1).ListIndex = i1
        'iterate values for vlookups
        For i2 = 1 To ColumnComments
            If Left(ArrSNRows(i1).ValuesList(i2), 10) = "=VLOOKUP(C" Then
                'alter vlookup reference to new list index position plus 2 header rows (i1 + 2)
                ArrSNRows(i1).ValuesList(i2) = Left(ArrSNRows(i1).ValuesList(i2), 10) + Str(i1 + 2) + Mid(ArrSNRows(i1).ValuesList(i2), InStr(1, ArrSNRows(i1).ValuesList(i2), ","))
                ArrSNRows(i1).ValuesList(i2) = Replace(ArrSNRows(i1).ValuesList(i2), "=VLOOKUP(C ", "=VLOOKUP(C") 'get rid of space that pop up out of nowhere
            End If
        Next i2
    Next i1

End Sub

Public Function MoveArrayEntry(ByRef arrayReordering() As Variant, ByVal indexMoving As Integer, ByVal indexDestination As Integer) As Variant
'This function will move the given indexed entry to the indexed destination given within the array given.
    
    Dim t() As Variant 'copy of array given minus the moving index
    Dim M() As Variant 'only the moving index
    Dim arrFinal() As Variant 'finalized array ready to return from function
    Dim aMax As Integer 'number of rows in array
    Dim bMax As Integer 'number of columns in array
    Dim a As Integer 'integer object to iterate through all rows in one entry of the given array
    Dim c As Integer 'integer to copy values into T
    
    
    'resize T to fit array given
    aMax = UBound(arrayReordering)
    ReDim t(1 To (aMax - 1)) 'one less entry with indexMoving taken out
    ReDim M(1 To 1) 'only the moving index
    ReDim arrFinal(1 To aMax) 'same size as entering array
    
    
    'copy array to T minus indexMoving
    c = 0 'initialize counter
    For a = 1 To aMax 'a is index for giving array for this loop
        
        'not indexMoving
        If a <> indexMoving Then
            c = c + 1 'increment c
            'object variable
            If IsObject(arrayReordering(a)) Then
                Set t(c) = arrayReordering(a)
            'regular variable
            Else
                t(c) = arrayReordering(a)
            End If
        
        'ignore indexMoving for T but put into M
        ElseIf a = indexMoving Then
            c = c 'do not increment c
            'object variable
            If IsObject(arrayReordering(a)) Then
                Set M(1) = arrayReordering(a)
            'regular variable
            Else
                M(1) = arrayReordering(a)
            End If
        
        End If
    Next a
    
    
    'finalize array for return
    c = 0 'initialize counter
    For a = 1 To aMax 'a is index for receiving array this time
        
        'not indexDestination
        If a <> indexDestination Then
            c = c + 1 'increment c
            'object variable
            If IsObject(t(c)) Then
                Set arrFinal(a) = t(c)
            'regular variable
            Else
                arrFinal(a) = t(c)
            End If
        
        'copy M into destination index
        ElseIf a = indexDestination Then
            c = c 'do not increment c
            'object variable
            If IsObject(M(1)) Then
                Set arrFinal(a) = M(1)
            'regular variable
            Else
                arrFinal(a) = M(1)
            End If
        
        End If
    Next a
    
    
    'return array
    MoveArrayEntry = arrFinal

End Function

Public Sub ClearSNRows()

    'activate sheet
    SheetFOPS.Activate

    'clear list
    SheetFOPS.Range(Cells(FirstListRow, 1), Cells(LastListRow, ColumnComments)).ClearContents
    SheetFOPS.Range(Cells(FirstListRow, 1), Cells(LastListRow, ColumnComments)).Interior.Color = xlNone
    SheetFOPS.Range(Cells(FirstListRow, 1), Cells(LastListRow, ColumnComments)).ClearComments

End Sub

Public Sub WriteSummary()

    Dim i As Integer 'iterator
    Dim sRows As Integer 'Number of Final Ops
    
    'activate sheet
    SheetSummary.Activate
    
    'initialize summary rows total
    sRows = 12 'one extra for FX Complete
    
    'clear summary
    SheetSummary.Range(Cells(2, 2), Cells((sRows + 1), 3)).ClearContents
    
    'iterate snrows
    For i = 1 To UBound(ArrSNRows)
        '24Ks
        If Left(ArrSNRows(i).ValuesList(1), 3) = "24K" Then
            'increment total (adding 1 for title row in summary)
            SheetSummary.Cells((1 + (ArrSNRows(i).WaterfallIndex)), 2).Value = (SheetSummary.Cells((1 + (ArrSNRows(i).WaterfallIndex)), 2).Value) + 1
        '30Ks
        ElseIf Left(ArrSNRows(i).ValuesList(1), 3) = "30K" Then
            'increment total (adding 1 for title row in summary)
            SheetSummary.Cells((1 + (ArrSNRows(i).WaterfallIndex)), 3).Value = (SheetSummary.Cells((1 + (ArrSNRows(i).WaterfallIndex)), 3).Value) + 1
        End If
    Next i

End Sub

Public Sub WriteSNRows()

    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    
    'activate sheet
    SheetFOPS.Activate
    
    'turn off alerts
    Application.DisplayAlerts = False
    
    'iterate list area
    For Each c In SheetFOPS.Range(Cells(FirstListRow, 1), Cells(LastListRow, 1))
    
        'iterate each SN Row object and write to list
        For i = 1 To ColumnComments
            'values
            c.Offset(0, (i - 1)).Value = ArrSNRows((c.Row - FirstListRow + 1)).ValuesList(i)
            'colors
            c.Offset(0, (i - 1)).Interior.Color = ArrSNRows((c.Row - FirstListRow + 1)).ColorsList(i)
            'comments
            If Not ArrSNRows((c.Row - FirstListRow + 1)).CommentsList(i) = "" Then
                c.Offset(0, (i - 1)).AddComment
                c.Offset(0, (i - 1)).Comment.Text ArrSNRows((c.Row - FirstListRow + 1)).CommentsList(i)
            End If
        Next i
        
        'Update List Indexes
        c.Offset(0, 1).Value = ArrSNRows((c.Row - FirstListRow + 1)).ListIndex
        
        'update loading bar
        lBar.UpdateLoadingBar "Writing SN Rows to table...", (c.Row - (FirstListRow - 1)), (LastListRow - (FirstListRow - 1))
    
    Next c
    
    'turn on alerts
    Application.DisplayAlerts = True

End Sub

Public Sub FixDaysAtPWA()

    Dim strFormula As String 'Formula for entry into cell
    Dim c As Range 'iterator
    
    For Each c In Range(Cells(FirstListRow, 9), Cells(LastListRow, 9))
        strFormula = "=IF(ISBLANK(J" & CStr(c.Row) & "),"""",IF(H" & CStr(c.Row) & "-J" & CStr(c.Row) & ">0,H" & CStr(c.Row) & "-J" & CStr(c.Row) & ",TODAY()-J" & CStr(c.Row) & "))"
        c.Formula = strFormula
    Next c

End Sub

Public Sub CreateBackupSheetCopy()
    
    'activate sheet
    SheetFOPS.Activate

    'create copy of SheetFOPS
    SheetFOPS.Copy Before:=SheetFOPS
    
    'assign copy of SheetFOPS
    Set SheetFOPS_BACKUP = ActiveSheet
    
    'reselect original sheet
    SheetFOPS.Activate

End Sub

Public Sub DeleteBackupSheetCopy()
    
    'activate sheet
    SheetFOPS.Activate

    Application.DisplayAlerts = False
    SheetFOPS_BACKUP.Delete
    Application.DisplayAlerts = True

End Sub

