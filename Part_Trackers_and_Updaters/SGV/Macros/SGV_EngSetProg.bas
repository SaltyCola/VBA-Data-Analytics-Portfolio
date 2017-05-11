Attribute VB_Name = "SGV_EngSetProg"
Public Sub EngineSetProgress()

    Dim c As Range 'range iteration object
    Dim i As Integer 'integer iteration object
    Dim j As Integer 'integer iteration object
    Dim first080Col As Integer 'first visible column of part number
    Dim first180Col As Integer 'first visible column of part number
    Dim first280Col As Integer 'first visible column of part number
    Dim first380Col As Integer 'first visible column of part number
    Dim first480Col As Integer 'first visible column of part number
    Dim EngSet1 As SGV_EngineSet
    Dim EngSet2 As SGV_EngineSet
    Dim EngSet3 As SGV_EngineSet
    Dim EngSet4 As SGV_EngineSet
    Dim EngSet5 As SGV_EngineSet
    Dim EngSet6 As SGV_EngineSet
    Dim EngSet7 As SGV_EngineSet
    Dim EngSet8 As SGV_EngineSet
    Dim EngSet9 As SGV_EngineSet
    Dim EngSet10 As SGV_EngineSet
    
    
    
    'Clear Sets-------------------------------------
    Worksheets("Engine Set Progress").Range("B3:DF18").ClearContents
    Worksheets("Engine Set Progress").Range("B3:DF18").Font.Bold = False
    Worksheets("Engine Set Progress").Range("B3:DF18").Font.Color = RGB(0, 0, 0)
    '-----------------------------------------------
    
    
    
    'find first visible columns ++++++++++++++++++++++++++++++++++++++++++++++++++++
        '080
        Worksheets("5319080").Activate
        For Each c In Range("13:13")
            If (c.Column > 2) And (c.EntireColumn.Hidden = False) Then
                first080Col = c.Column
                Exit For
            End If
        Next c
        '180
        Worksheets("5319180").Activate
        For Each c In Range("13:13")
            If (c.Column > 2) And (c.EntireColumn.Hidden = False) Then
                first180Col = c.Column
                Exit For
            End If
        Next c
        '280
        Worksheets("5319280").Activate
        For Each c In Range("13:13")
            If (c.Column > 2) And (c.EntireColumn.Hidden = False) Then
                first280Col = c.Column
                Exit For
            End If
        Next c
        '380
        Worksheets("5319380").Activate
        For Each c In Range("13:13")
            If (c.Column > 2) And (c.EntireColumn.Hidden = False) Then
                first380Col = c.Column
                Exit For
            End If
        Next c
        '480
        Worksheets("5319480").Activate
        For Each c In Range("13:13")
            If (c.Column > 2) And (c.EntireColumn.Hidden = False) Then
                first480Col = c.Column
                Exit For
            End If
        Next c
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



    'initialize class objects -------------------------------------------------------
    Set EngSet1 = New SGV_EngineSet
    EngSet1.SetNum = 1
    Set EngSet2 = New SGV_EngineSet
    EngSet2.SetNum = 2
    Set EngSet3 = New SGV_EngineSet
    EngSet3.SetNum = 3
    Set EngSet4 = New SGV_EngineSet
    EngSet4.SetNum = 4
    Set EngSet5 = New SGV_EngineSet
    EngSet5.SetNum = 5
    Set EngSet6 = New SGV_EngineSet
    EngSet6.SetNum = 6
    Set EngSet7 = New SGV_EngineSet
    EngSet7.SetNum = 7
    Set EngSet8 = New SGV_EngineSet
    EngSet8.SetNum = 8
    Set EngSet9 = New SGV_EngineSet
    EngSet9.SetNum = 9
    Set EngSet10 = New SGV_EngineSet
    EngSet10.SetNum = 10
    '--------------------------------------------------------------------------------
    
    
    
    'set starting and stopping points=============================================
        'engine set 1
        With EngSet1
            '080
            .sp080 = first080Col
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = first180Col
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = first280Col
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = first380Col
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = first480Col
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 2
        With EngSet2
            '080
            .sp080 = EngSet1.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet1.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet1.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet1.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet1.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 3
        With EngSet3
            '080
            .sp080 = EngSet2.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet2.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet2.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet2.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet2.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 4
        With EngSet4
            '080
            .sp080 = EngSet3.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet3.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet3.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet3.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet3.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 5
        With EngSet5
            '080
            .sp080 = EngSet4.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet4.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet4.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet4.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet4.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 6
        With EngSet6
            '080
            .sp080 = EngSet5.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet5.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet5.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet5.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet5.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 7
        With EngSet7
            '080
            .sp080 = EngSet6.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet6.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet6.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet6.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet6.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 8
        With EngSet8
            '080
            .sp080 = EngSet7.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet7.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet7.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet7.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet7.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 9
        With EngSet9
            '080
            .sp080 = EngSet8.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet8.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet8.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet8.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet8.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
        'engine set 10
        With EngSet10
            '080
            .sp080 = EngSet9.ep080 + 1
            .ep080 = .sp080 + (14 - 1)
            
            '180
            .sp180 = EngSet9.ep180 + 1
            .ep180 = .sp180 + (11 - 1)
            
            '280
            .sp280 = EngSet9.ep280 + 1
            .ep280 = .sp280 + (12 - 1)
            
            '380
            .sp380 = EngSet9.ep380 + 1
            .ep380 = .sp380 + (8 - 1)
            
            '480
            .sp480 = EngSet9.ep480 + 1
            .ep480 = .sp480 + (9 - 1)
            
        End With
    '=============================================================================
    
    
    
    'read SNs into arrays##########################################################
        
        'engine set 1
        With EngSet1
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 2
        With EngSet2
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 3
        With EngSet3
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 4
        With EngSet4
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 5
        With EngSet5
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 6
        With EngSet6
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 7
        With EngSet7
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 8
        With EngSet8
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 9
        With EngSet9
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        'engine set 10
        With EngSet10
            '080
            For i = .sp080 To .ep080 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319080").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr080.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '180
            For i = .sp180 To .ep180 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319180").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr180.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '280
            For i = .sp280 To .ep280 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319280").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr280.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '380
            For i = .sp380 To .ep380 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319380").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr380.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
            '480
            For i = .sp480 To .ep480 'columns iteration
                'rows iteration
                For j = 20 To 35
                    'SN position found
                    If Not Worksheets("5319480").Cells(j, i).Interior.Color = RGB(255, 255, 255) Then
                        'increment array value
                        .arr480.OpRowIncrement (j)
                        Exit For
                    End If
                Next j
            Next i
        End With
        
    '##############################################################################



    ' print to engine set sheet ---------------------------------------------------
    
    EngSet1.PrintEngSet
    EngSet2.PrintEngSet
    EngSet3.PrintEngSet
    EngSet4.PrintEngSet
    EngSet5.PrintEngSet
    EngSet6.PrintEngSet
    EngSet7.PrintEngSet
    EngSet8.PrintEngSet
    EngSet9.PrintEngSet
    EngSet10.PrintEngSet
    
    '------------------------------------------------------------------------------



End Sub
