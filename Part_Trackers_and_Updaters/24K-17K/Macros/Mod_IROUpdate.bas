Attribute VB_Name = "Mod_IROUpdate"
Sub IROUpdate()

    Dim lBar As TC_LoadingBar 'animated loading bar
    Dim numCols As Integer 'number of SN Columns being updated
    Dim iCol As Integer 'current SN column being updated
    Dim arrSerialNums() As Variant 'array holding the list of SNs to be updated
    Dim wrkb As Workbook 'current tracker file workbook
    Dim wrks As Worksheet 'current tracker worksheet
    Dim wrkbIRO As Workbook 'IRO Log workbook
    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    
    'initialize loading bar
    Set lBar = New TC_LoadingBar
    
    'initialize counters
    numCols = 0
    
    'get current workbook and worksheet
    Set wrkb = ActiveWorkbook
    Set wrks = wrkb.ActiveSheet
    
    'grab number of columns
    For Each c In wrks.Range("5:5")
        If c.Column > 4 And Not IsEmpty(c) Then
            'increment counter
            numCols = numCols + 1
            'add SN to array
            ReDim Preserve arrSerialNums(1 To 3, 1 To numCols) As Variant
            arrSerialNums(1, numCols) = c.Value
            arrSerialNums(2, numCols) = c.Address
            arrSerialNums(3, numCols) = xlNone 'initialize color to none
        'ignore empty redline cells, but exit on first white/nocolor empty cell
        ElseIf c.Column > 4 And IsEmpty(c) And IsEmpty(c.Offset(0, 1)) And (Range(c, c.Offset(0, 1)).EntireColumn.Interior.Color = RGB(255, 255, 255) Or Range(c, c.Offset(0, 1)).EntireColumn.Interior.Color = xlNone) Then
            Exit For
        End If
    Next c
    
    'open IRO Log workbook in READ ONLY mode =================================================
    Workbooks.Open "\\PUSLMA03\Operations\5. Cell Data\PWAA_GTF\FBC_Continuous Improvement\Fan Blade DIVE\IRO_16194_Log.xlsm", , True
    Set wrkbIRO = ActiveWorkbook
    '=========================================================================================
    
    'activate tracker worksheet
    wrks.Activate
    
    'Grab Color Indexes and write them into array
    For i = 1 To UBound(arrSerialNums, 2)
        
        For Each c In wrkbIRO.Worksheets("Data Sheet").Range("C:C")
            'empty SN cell found, exit this for loop
            If IsEmpty(c) Then
                Exit For
            'SN found in IRO Log
            ElseIf arrSerialNums(1, i) = c.Value Then
                
                If c.Offset(0, 8) = "Fail" And c.Offset(0, 23) = "Reject" Then
                    arrSerialNums(3, i) = RGB(255, 0, 0)
                ElseIf c.Offset(0, 8) = "Fail" And c.Offset(0, 23) = "" Then
                    arrSerialNums(3, i) = RGB(255, 255, 0)
                ElseIf c.Offset(0, 8) = "Fail" And c.Offset(0, 23) = "NQM" Then
                    arrSerialNums(3, i) = RGB(0, 255, 0)
                ElseIf c.Offset(0, 8) = "Fail" And c.Offset(0, 23) = "Accept" Then
                    arrSerialNums(3, i) = RGB(0, 255, 0)
                ElseIf c.Offset(0, 8) = "Fail" And c.Offset(0, 23) = "Reinspect" Then
                    arrSerialNums(3, i) = RGB(255, 102, 0)
                ElseIf c.Offset(0, 8) = "Pass" Then
                    arrSerialNums(3, i) = RGB(0, 255, 0)
                ElseIf c.Offset(0, 8) = "Not Completed" Then
                    arrSerialNums(3, i) = RGB(0, 0, 255)
                End If
                
            End If
        Next c
        
        'write color to Tracker
        wrks.Range(arrSerialNums(2, i)).Offset(-1, 0).Interior.Color = arrSerialNums(3, i)
        
        'call loading bar updater
        lBar.UpdateLoadingBar " Updating from IRO Log...", i, numCols
        
    Next i
    
    'close IRO Log workbook without saving ===================================================
    Workbooks("IRO_16194_Log").Close False
    '=========================================================================================

End Sub

