Attribute VB_Name = "SGV_Miscellaneous"
Sub CellValuePropertiesTest()

    Dim tCell As Range
    
    Set tCell = Range("HG29")

    MsgBox ".Formula: " & tCell.Formula & vbNewLine & ".Text: " & tCell.Text & vbNewLine & ".Value: " & tCell.Value & vbNewLine & ".Value2: " & tCell.Value2

End Sub

Sub CellFormulaTest()

    Dim sCell As Range
    Dim dCell As Range
    
    Set sCell = Range("HG29")
    Set dCell = Range("HG49")
    
    sCell.Formula = dCell.Formula

    'So for transferring over all SN columns into current version, use .Formula as this
    ' will ensure all cell values transfer acurrately. (with no #REF! errors)
    
End Sub

Public Sub GrabSnColumns()

    Dim i As Double 'Counter for accessing each part number's WIP tab
    Dim sn As Double 'Counter for naming SN_Column objects
    Dim tb As Worksheet 'Iteration worksheet object for finding part number WIP tabs
    Dim rngSNRow As Range 'Iteration range object for finding S/N row number within each tab
    Dim rwSN As Double 'Row number for the current tab's S/N row
    Dim rngSNCell As Range 'Iteration range object for grabbing each SN col
    Dim coll080s As Collection 'Collection of 080 SN_Column objects
    Dim coll180s As Collection 'Collection of 180 SN_Column objects
    Dim coll280s As Collection 'Collection of 280 SN_Column objects
    Dim coll380s As Collection 'Collection of 380 SN_Column objects
    Dim coll480s As Collection 'Collection of 480 SN_Column objects
    Dim CurrentSN As SGV_Column 'SGV_Column object for adding to corresponding collection
    Dim c As Double 'Generic iteration counter
    Dim tempQnArray() As Variant 'Array for storing qn info
    Dim tempOpDateArray() As Variant 'Array for storing op date info
    Dim tempNotesArray() As Variant 'Array for storing notes info
    
    'initialize collections
    Set coll080s = New Collection
    Set coll180s = New Collection
    Set coll280s = New Collection
    Set coll380s = New Collection
    Set coll480s = New Collection
    
    'activate source workbook
    Workbooks("Copy of v6").Activate
    
    'initialize i counter
    i = 0
    
    'iterate all tabs
    For Each tb In ActiveWorkbook.Worksheets
    
        'turn off screen updating
        Application.ScreenUpdating = False
        
        'part number WIP tab found
        If tb.Name = "5319" & i & "80" Then

            'initialize sn counter
            sn = 1
        
            'find S/N row
            For Each rngSNRow In tb.Range("B:B")
                If rngSNRow.Value = "S/N" Then: rwSN = rngSNRow.Row
            Next rngSNRow
            
            'iterate S/N row
            For Each rngSNCell In tb.Range(rwSN & ":" & rwSN)
                
                'end iteration if SN title cell empty
                If (rngSNCell.Column > 2) And Not (rngSNCell.EntireColumn.Hidden) And (IsEmpty(rngSNCell)) Then
                    Exit For
                'if active SN column found
                ElseIf (rngSNCell.Column > 2) And Not (rngSNCell.EntireColumn.Hidden) And Not (IsEmpty(rngSNCell)) Then
                                        
                    'Reset current objects
                    Set CurrentSN = Nothing
                    ReDim tempQnArray(1 To 6, 1 To 2)
                    ReDim tempOpDateArray(1 To 17, 1 To 2)
                    ReDim tempNotesArray(1 To 4, 1 To 2)
                                        
                    
'set current SN_Column object properties =================================================
                    Set CurrentSN = New SGV_Column
                    CurrentSN.PartNumber = rngSNCell.Offset(-1, 0).Value
                    CurrentSN.SerialNumber = rngSNCell.Value
                    For c = 1 To 6
                        tempQnArray(c, 1) = rngSNCell.Offset(c, 0).Value
                        tempQnArray(c, 2) = rngSNCell.Offset(c, 0).Interior.Color
                    Next c
                    CurrentSN.QnArray = tempQnArray
                    For c = 1 To 17
                        tempOpDateArray(c, 1) = rngSNCell.Offset((6 + c), 0).Text
                        tempOpDateArray(c, 2) = rngSNCell.Offset((6 + c), 0).Interior.Color
                    Next c
                    CurrentSN.OpDateArray = tempOpDateArray
                    For c = 1 To 4
                        tempNotesArray(c, 1) = rngSNCell.Offset((23 + c), 0).Value
                        tempNotesArray(c, 2) = rngSNCell.Offset((23 + c), 0).Interior.Color
                    Next c
                    CurrentSN.NotesArray = tempNotesArray
'set current SN_Column object properties =================================================
                    
                                    
                    'add currentSN to corresponding collection
                    If tb.Name = "5319080" Then
                        coll080s.Add CurrentSN, Str(sn)
                    ElseIf tb.Name = "5319180" Then
                        coll180s.Add CurrentSN, Str(sn)
                    ElseIf tb.Name = "5319280" Then
                        coll280s.Add CurrentSN, Str(sn)
                    ElseIf tb.Name = "5319380" Then
                        coll380s.Add CurrentSN, Str(sn)
                    ElseIf tb.Name = "5319480" Then
                        coll480s.Add CurrentSN, Str(sn)
                    End If
                    
                    'increment sn counter
                    sn = sn + 1
                    
                End If
                
            Next rngSNCell
        
            'increment i counter
            If i < 4 Then: i = i + 1
        
        End If
        
        'turn on screen updating
        Application.ScreenUpdating = True
        
    Next tb
    
    'Call writing function
    Call WriteToCurrentVersion(coll080s, coll180s, coll280s, coll380s, coll480s)

End Sub

Public Sub WriteToCurrentVersion(ByVal coll080s As Collection, ByVal coll180s As Collection, ByVal coll280s As Collection, ByVal coll380s As Collection, ByVal coll480s As Collection)

    Dim i As Double 'Counter for accessing each part number's WIP tab
    Dim tb As Worksheet 'Iteration worksheet object for finding part number WIP tabs
    Dim rngSNRow As Range 'Iteration range object for finding S/N row number within each tab
    Dim rwSN As Double 'Row number for the current tab's S/N row
    Dim rngSNCell As Range 'Iteration range object for writing to each SN col
    Dim sgv As SGV_Column 'Iteration SGV_Column object for grabbing each sgv in a collection
    Dim currColl As Collection 'Current collection to pull from
    Dim z As Range 'Generic Range iteration object
    Dim zz As Double 'Generic iteration counter
    
    'activate target workbook
    Workbooks("PWAA Lansing NGPF Vanes WIP Status and Detail Tracking_JMF Planning v7").Activate
    
    'initialize i counter
    i = 0
    
    'iterate all tabs
    For Each tb In ActiveWorkbook.Worksheets
    
        'turn off screen updating
        Application.ScreenUpdating = False
        
        'part number WIP tab found
        If tb.Name = "5319" & i & "80" Then
            
            'get correct collection to write from
            If tb.Name = "5319080" Then
                Set currColl = coll080s
            ElseIf tb.Name = "5319180" Then
                Set currColl = coll180s
            ElseIf tb.Name = "5319280" Then
                Set currColl = coll280s
            ElseIf tb.Name = "5319380" Then
                Set currColl = coll380s
            ElseIf tb.Name = "5319480" Then
                Set currColl = coll480s
            End If
            
            'find S/N row
            For Each rngSNRow In tb.Range("B:B")
                If rngSNRow.Value = "S/N" Then: rwSN = rngSNRow.Row
            Next rngSNRow
            
            'iterate S/N row
            For Each rngSNCell In tb.Range(rwSN & ":" & rwSN)
            
                'search corresponding collection for correct SGV_Column object
                For Each sgv In currColl
                    If rngSNCell.Value = sgv.SerialNumber Then
                        
                        'write sgv contents to rngSNCell's column
                        zz = 1
                        For Each z In Range(rngSNCell.Offset(7, 0), rngSNCell.Offset(23, 0))
                            z.Value = sgv.OpDateArray(zz, 1)
                            z.Interior.Color = sgv.OpDateArray(zz, 2)
                            zz = zz + 1
                        Next z
                        zz = 1
                        For Each z In Range(rngSNCell.Offset(24, 0), rngSNCell.Offset(27, 0))
                            z.Value = sgv.NotesArray(zz, 1)
                            z.Interior.Color = sgv.NotesArray(zz, 2)
                            zz = zz + 1
                        Next z
                        
                    End If
                Next sgv
            
            Next rngSNCell
        
            'increment i counter
            If i < 4 Then: i = i + 1
            
        End If
        
    'turn on screen updating
        Application.ScreenUpdating = True
        
    Next tb

End Sub

Public Sub WritingTest(ByVal coll080s As Collection, ByVal coll180s As Collection, ByVal coll280s As Collection, ByVal coll380s As Collection, ByVal coll480s As Collection)

'DEBUGGING====================================================================================
    Dim z As Double
    Dim y As Variant
    Dim x As Double
    Dim yStr As String
    
    For z = 1 To 10
        
        MsgBox coll080s.Item(z).PartNumber & "     " & coll080s.Item(z).SerialNumber
        
        y = coll080s.Item(z).QnArray
        
        For x = 1 To 6
            yStr = yStr + (y(x, 1) & ":   " & y(x, 2) & vbNewLine)
        Next x
        
        MsgBox yStr
        
        yStr = ""
        
    Next z
'DEBUGGING====================================================================================

End Sub

Public Sub MakeWhiteCells()

    Dim c As Range
    
    For Each c In Range("JJ20:NC35")
    
        If Not (c.Interior.Color = RGB(146, 208, 80)) And Not (c.Interior.Color = RGB(247, 150, 70)) And Not (c.Interior.Color = RGB(255, 255, 0)) And Not (c.Interior.Color = RGB(250, 191, 143)) Then
        
            Application.Goto c
            c.Interior.Color = xlNone
        
        End If
    Next c

End Sub

