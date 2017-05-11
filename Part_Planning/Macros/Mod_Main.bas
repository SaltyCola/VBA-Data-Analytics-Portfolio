Attribute VB_Name = "Mod_Main"
'===================Public Variables===================='
Public CurrentDateCell As Range
Public firstRectangle As Boolean
Public rctLeft As Double
Public rctTop As Double
Public rctWidth As Double
Public rctHeight As Double
Public BottomLineRectangle As Boolean
Public FinalVerticalRectangle As Boolean

Public frmEngSetUpdater As ES_EngineSetUpdater
Public esCol As Integer
Public esRow As Integer
Public boolAllowToggles As Boolean
Public boolToggled As Boolean
Public TodayCell As Range

Public clrWhite As Long
Public clrGreen As Long
Public clrYellow As Long
Public clrRed As Long
'===================Public Variables===================='

Public Sub FindCurrentDateCells()

    
    '=================Delete All Shapes================='
    ActiveSheet.DrawingObjects.Delete
    '=================Delete All Shapes================='
    
    
    Dim rngTopRow As Range
    Dim col As Range
    Dim row As Integer
    Dim today As Date
    
    'define colors
    clrWhite = RGB(255, 255, 255)
    clrGreen = RGB(146, 208, 80)
    clrYellow = RGB(255, 255, 0)
    clrRed = RGB(255, 0, 0)
    
    'define today
    today = Date
    
    'set firstRectangle to True
    firstRectangle = True
    
    'set BottomLineRectangle to False
    BottomLineRectangle = False
    
    'set FinalVerticalRectangle to true
    FinalVerticalRectangle = True
    
    'define rngSrchRow
    Set rngTopRow = Worksheets("NEO 5322121 Aggressive LTs").Range("6:6")
    
    'iterate top row for columns
    For Each col In rngTopRow.SpecialCells(xlCellTypeVisible)
        
        'exit sub on black line found
        If col.Interior.Color = RGB(0, 0, 0) Then
            'go to engine set updater form
            GoTo lineCallESU
        End If
        
        'ignore first three columns
        If col.Column > 3 Then
    
            'iterate row
            For row = 7 To 33
                
                'finish line on bottom right corner
                If IsEmpty(Worksheets("NEO 5322121 Aggressive LTs").Cells(row, col.Column)) Then
                    BottomLineRectangle = True
                    Set CurrentDateCell = Worksheets("NEO 5322121 Aggressive LTs").Cells(row, col.Column)
                    Call AddShapeToCell
                    GoTo lineNextcol
                'find first date <= today
                ElseIf DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells(row, col.Column).Value) <= today Then
                    'current date cell
                    Set CurrentDateCell = Worksheets("NEO 5322121 Aggressive LTs").Cells(row, col.Column)
                    Call AddShapeToCell
                    GoTo lineNextcol
                End If
            
            Next row
        End If
lineNextcol:
    Next col


lineCallESU:
    Call EngineSetUpdater


End Sub

Public Sub AddShapeToCell()

    Dim cllLeft As Double
    Dim cllTop As Double
    Dim cllWidth As Double
    Dim cllHeight As Double
    Dim shpRectangleH As Shape
    Dim shpRectangleV As Shape
    
    'define cll dimensions
    cllLeft = CurrentDateCell.Left
    cllTop = CurrentDateCell.Top
    cllWidth = CurrentDateCell.Width
    cllHeight = CurrentDateCell.Height
    
    
    'if not bottom line rectangle
    If Not BottomLineRectangle Then
    
        'create and place vertical
        If (firstRectangle = False) And ((cllTop - rctTop + 1) >= 0) Then
            Set shpRectangleV = ActiveSheet.Shapes.AddShape(msoShapeRectangle, (rctLeft + rctWidth - 1), rctTop, 2, (cllTop - rctTop + 1))
        ElseIf (firstRectangle = False) And ((cllTop - rctTop + 1) < 0) Then
            Set shpRectangleV = ActiveSheet.Shapes.AddShape(msoShapeRectangle, (rctLeft + rctWidth - 1), cllTop, 2, (rctTop - cllTop + 1))
        End If
        
        'create and place horrizontal shape
        Set shpRectangleH = ActiveSheet.Shapes.AddShape(msoShapeRectangle, cllLeft, cllTop, cllWidth, 2)
        rctLeft = shpRectangleH.Left
        rctTop = shpRectangleH.Top
        rctWidth = shpRectangleH.Width
        rctHeight = shpRectangleH.Height
        
        'set firstRectangle to false
        firstRectangle = False
    
    
    'Bottom Line Reached
    ElseIf BottomLineRectangle Then
        
        'place final vertical
        If (FinalVerticalRectangle) And ((cllTop - rctTop + 1) >= 0) Then
            Set shpRectangleV = ActiveSheet.Shapes.AddShape(msoShapeRectangle, (rctLeft + rctWidth - 1), rctTop, 2, (cllTop - rctTop + 1))
        ElseIf (FinalVerticalRectangle) And ((cllTop - rctTop + 1) < 0) Then
            Set shpRectangleV = ActiveSheet.Shapes.AddShape(msoShapeRectangle, (rctLeft + rctWidth - 1), cllTop, 2, (rctTop - cllTop + 1))
        End If
        
        'create and place horrizontal shape
        Set shpRectangleH = ActiveSheet.Shapes.AddShape(msoShapeRectangle, cllLeft, cllTop, cllWidth, 2)
        
        'set FinalVerticalRectangle to false
        FinalVerticalRectangle = False
        
    End If
    
    
End Sub

Public Sub EngineSetUpdater()
    
    'initialize booleans
    boolToggled = False
    
    'Load Engine Set Updater
    Set frmEngSetUpdater = New ES_EngineSetUpdater
    
    'initialize button captions
    Dim c As Double
    For c = 7 To 33
        frmEngSetUpdater.Controls("ToggleR" & c).Caption = Worksheets("NEO 5322121 Aggressive LTs").Cells(c, 3)
    Next c
    
    'set focus and show form
    frmEngSetUpdater.TextBoxESU.SetFocus
    frmEngSetUpdater.Show

End Sub

Public Sub ToggleButtonHandler()

    'define integer for recoloring columns
    Dim i As Integer

    'find today cell
    Dim tcRng As Range
    Dim tc As Range
    Set tcRng = Worksheets("NEO 5322121 Aggressive LTs").Range(Cells(7, esCol), Cells(32, esCol))
    For Each tc In tcRng
        If tc.Value <= Date Then
            Set TodayCell = tc
            Exit For
        End If
    Next tc


    'if button is toggled
    If boolToggled Then
    
        'green fill (at or above day line)
        If (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Value) >= Date) Then
            For i = esRow To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Interior.Color = clrGreen
                frmEngSetUpdater.Controls("ToggleR" & esRow).BackColor = clrGreen
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than esRow
                If Not i = esRow Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).Enabled = False
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        'yellow fill (up to 3 days below line)
        ElseIf (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Value) < Date) And (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Value) >= (Date - 3)) Then
            For i = esRow To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Interior.Color = clrYellow
                frmEngSetUpdater.Controls("ToggleR" & esRow).BackColor = clrYellow
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than esRow
                If Not i = esRow Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).Enabled = False
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        'red fill (more than 3 days below line)
        ElseIf (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Value) < (Date - 3)) Then
            For i = esRow To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Interior.Color = clrRed
                frmEngSetUpdater.Controls("ToggleR" & esRow).BackColor = clrRed
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than esRow
                If Not i = esRow Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).Enabled = False
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        End If
        
    'if button is untoggled
    ElseIf Not boolToggled Then
    
        'white out button and cell, and enabled first button below
        Worksheets("NEO 5322121 Aggressive LTs").Cells(esRow, esCol).Interior.Color = clrWhite
        frmEngSetUpdater.Controls("ToggleR" & esRow).BackColor = clrWhite
        'enable first button below except for last row
        If esRow < 33 Then
            frmEngSetUpdater.Controls("ToggleR" & (esRow + 1)).Enabled = True
        End If
        'exit sub if last row button untoggled
        If esRow = 33 Then
            Exit Sub
        End If
        
        'green fill (at or above day line)
        If (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Value) >= Date) Then
            For i = (esRow + 1) To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Interior.Color = clrGreen
                frmEngSetUpdater.Controls("ToggleR" & (esRow + 1)).BackColor = clrGreen
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than (esRow + 1)
                If Not i = (esRow + 1) Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        'yellow fill (up to 3 days below line)
        ElseIf (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Value) < Date) And (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Value) >= (Date - 3)) Then
            For i = (esRow + 1) To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Interior.Color = clrYellow
                frmEngSetUpdater.Controls("ToggleR" & (esRow + 1)).BackColor = clrYellow
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than (esRow + 1)
                If Not i = (esRow + 1) Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        'red fill (more than 3 days below line)
        ElseIf (DateValue(Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Value) < (Date - 3)) Then
            For i = (esRow + 1) To 33
                boolAllowToggles = False
                'for i = esRow
                Worksheets("NEO 5322121 Aggressive LTs").Cells((esRow + 1), esCol).Interior.Color = clrRed
                frmEngSetUpdater.Controls("ToggleR" & (esRow + 1)).BackColor = clrRed
                frmEngSetUpdater.Controls("ToggleR" & i).Value = True
                'for all i greater than (esRow + 1)
                If Not i = (esRow + 1) Then
                    Worksheets("NEO 5322121 Aggressive LTs").Cells(i, esCol).Interior.Color = clrGreen
                    frmEngSetUpdater.Controls("ToggleR" & i).BackColor = clrGreen
                End If
            Next i
            boolAllowToggles = True
        End If
        
    End If
        
        
End Sub

Sub alsiudhflashdlfjhalsdkjfhlkajsdhfqowiepowepojwifdpwppwpefpjwepwepjdpiwdgpiwdpignpdiwvpniwpnidpi2389u9823y2394239842803u0823j4i23()

    '=================Delete All Shapes================='
    ActiveSheet.DrawingObjects.Delete
    '=================Delete All Shapes================='

End Sub
