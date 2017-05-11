Attribute VB_Name = "SGV_Summary"
'Double Click Event Code is located in the "Sheet12 (WIP Summary)" Code

Public Sub ClearSummary()
    
    'turn off screen updating
    Application.ScreenUpdating = False
    
    'activate WIP Summary Tab
    Worksheets("WIP Summary").Activate
    
    'clear all summary cells
    Worksheets("WIP Summary").Range("B5:R409").ClearContents
    Worksheets("WIP Summary").Range("B5:R409").Interior.Color = xlNone
    
    'turn on screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub UpdateSummary()

    'clear all cells
    Call ClearSummary

    'declare variables (SGV_List object declared within loops below)
    Dim transArray() As String 'Transposing array for sending to SGV_List Object
    Dim prtCll As Range 'range iteration object for part numbers in Summary Tab
    Dim opCll As Range 'range iteration object for op names in Summary Tab
    Dim pn As String 'current part number
    Dim Op As String 'current op name
    Dim wrk As Worksheet 'current search worksheet object
    Dim snRow As Double 'current wrk's "S/N" row
    Dim searchColor As Long 'current SN color searching for (green or yellow)
    Dim search2Color As Long 'current SN color searching for (salmon or orange)
    Dim searchRow As Double 'current row searching through in wrk
    Dim sr As Range 'range iteration object for finding current searchRow
    Dim srCll As Range 'range iteration object for storing SN's in transArray per current searchRow
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'activate WIP Summary Tab
    Worksheets("WIP Summary").Activate
    Application.ScreenUpdating = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'iterate part numbers
    For Each prtCll In Worksheets("WIP Summary").Range(Cells(5, 1), Cells(9, 1))
        
        'grab part number
        pn = prtCll.Value
        
        'iterate ops (skips "repair cell" column)
        For Each opCll In Worksheets("WIP Summary").Range(Cells(4, 3), Cells(4, 18))
            
            'reset searchColors to green
            searchColor = RGB(146, 208, 80) 'Green
            search2Color = RGB(250, 191, 143) 'Salmon
            
            'reset transArray
            ReDim transArray(0 To 0)
            
            'grab op name
            Op = opCll.Value
            
            'alter op name and search color for Fountain Plating Op Columns
            If Op = "Back From Fountain Plating" Then
                Op = "Fountain Plating  -  IHC"
                searchColor = RGB(146, 208, 80) 'Green
                search2Color = RGB(250, 191, 143) 'Salmon
            ElseIf Op = "Hardcoat Outsource" Then
                Op = "Fountain Plating  -  IHC"
                searchColor = RGB(255, 255, 0) 'Yellow
                search2Color = RGB(247, 150, 70) 'Orange
            End If
            
            'set current search tab
            Set wrk = Worksheets(pn)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'activate current worksheet
            Application.ScreenUpdating = False
            wrk.Activate
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            'find S/N row and search row
            For Each sr In wrk.Range(Cells(10, 2), Cells(40, 2))
                'S/N row found
                If sr.Value = "S/N" Then
                    snRow = sr.Row
                'search row found
                ElseIf sr.Value = Op Then
                    searchRow = sr.Row
                    Exit For
                End If
            Next sr
            
            'search through row for SN's
            For Each srCll In wrk.Range(searchRow & ":" & searchRow)
                'exit searching through row if SN cell is blank
                If (srCll.Column > 2) And Not (srCll.EntireColumn.Hidden) And (IsEmpty(wrk.Cells(snRow, srCll.Column))) Then
                    Exit For
                'SN found via searchColor
                ElseIf (srCll.Column > 2) And Not (srCll.EntireColumn.Hidden) And ((srCll.Interior.Color = searchColor) Or (srCll.Interior.Color = search2Color)) And (Range(srCll.Offset(-1, 0), Cells(20, srCll.Column)).Interior.Color = RGB(255, 255, 255)) Then
                    'append 2 new object spaces to transArray
                    If UBound(transArray) = 0 Then
                        ReDim transArray(1 To (UBound(transArray) + 2))
                    ElseIf UBound(transArray) > 0 Then
                        ReDim Preserve transArray(1 To (UBound(transArray) + 2))
                    End If
                    'add SN to transArray
                    transArray(UBound(transArray) - 1) = Right(wrk.Cells(snRow, srCll.Column).Value, 5)
                    'add op date to transArray
                    transArray(UBound(transArray)) = srCll.Text
                End If
            Next srCll
            
            'declare SGV_List Object
            Dim sgvList As SGV_List
            'create SGV_List Object
            Set sgvList = New SGV_List
            'set SGV_List properties
            sgvList.Color = prtCll.Interior.Color
            sgvList.Op = opCll.Value
            sgvList.PartNum = prtCll.Value
            sgvList.SGVArray = transArray
            sgvList.Total = UBound(transArray) / 2
            'run SGV_List Methods
            sgvList.PrintTotal
            sgvList.PrintList
            'Erase sgvList
            Set sgvList = Nothing
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'activate summary worksheet
            Worksheets("WIP Summary").Activate
            Application.ScreenUpdating = True
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Next opCll
    Next prtCll
    
    'calculate totals
    Calculate

End Sub
