VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_InfoPage 
   Caption         =   "SN_InfoPage"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   OleObjectBlob   =   "SN_InfoPage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_InfoPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Userform_Activate()

    'disable toggles subs
    boolTogsAllowed = False

    'Date/Color TextBox Fill
    Dim dateCell As Range
    Dim snRange As Range
    Dim rowIP As Double
    
    'Set Serial Number Column
    snSearchCol = snTitleCell.Column
    
    Set snRange = Worksheets("NEO 5322121").Range(Cells(7, snSearchCol), Cells(43, snSearchCol))
    'iterate SN column
    For Each dateCell In snRange
    
        'skip rows 34 through 37
        If (dateCell.Row = 34) Or (dateCell.Row = 35) Or (dateCell.Row = 36) Or (dateCell.Row = 37) Then
            GoTo lineNextDateCell
        End If
        
        '================================================================================='
        'initialize boolFlagCell and boolFlagCellClear
        boolFlagCell = False
        boolFlagCellClear = False
        'search for errors
        If (IsError(dateCell.Value)) Then
            'Load SN_DeleteError form
            Set frmSNDeleteError = New SN_DeleteError
            boolCanceled = False
            'initialize txtboxes
            frmSNDeleteError.TextBox1.Value = dateCell.Address(False, False)
            'define error cell range and go to error cell
            rngErrorCell = dateCell.Address
            Application.Goto Worksheets("NEO 5322121").Range(rngErrorCell)
            'show form
            frmSNDeleteError.Show
            'if SN_DeleteError is exitted
            If boolCanceled Then
                Exit Sub
            End If
        End If
        '================================================================================='
        
        'tracker cell has color
        If Not (dateCell.Interior.Color = clrBlank) Then
            
            'for SN rows 7 through 33
            If (dateCell.Row >= 7) And (dateCell.Row <= 33) Then
                rowIP = dateCell.Row - 6
            'for SN rows 38 through 43
            ElseIf (dateCell.Row >= 38) And (dateCell.Row <= 43) Then
                rowIP = dateCell.Row - 10
            End If
            
        'Fill corresponding info page txtbox
            'txtbox color
            Me.Controls("TextBox" & "R" & rowIP).BackColor = dateCell.Interior.Color
            
            'not flagged as error cell
            If Not (boolFlagCell) Then
                'filled date
                If Not (dateCell.Value = "") And Len(Str(dateCell.Value)) > 5 Then
                    Me.Controls("TextBox" & "R" & rowIP).Value = Left(Str(dateCell.Value), Len(Str(dateCell.Value)) - 5)
                ElseIf Not (dateCell.Value = "") And Len(Str(dateCell.Value)) <= 5 Then
                    Me.Controls("TextBox" & "R" & rowIP).Value = Str(dateCell.Value)
                'empty date
                ElseIf (dateCell.Value = "") Then
                    Me.Controls("TextBox" & "R" & rowIP).Value = ""
                End If
                
            '=================================================================================='
            'flagged as error cell
            ElseIf (boolFlagCell) Then
                Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(0, 0, 0)
                Me.Controls("TextBox" & "R" & rowIP).Value = "FLAGGED AS ERROR"
                Me.Controls("TextBox" & "R" & rowIP).ForeColor = RGB(192, 0, 0)
                'disable and lock entire row
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                    'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Enabled = False
                End If
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
            End If
            '=================================================================================='
            
            
            
            'disable and lock row
            'green
            If Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(146, 208, 80) Then
                Me.Controls("ToggleButton" & "1R" & rowIP).Value = True
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Enabled = False
                End If
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'dark green
            ElseIf Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(79, 98, 40) Then
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Value = True
                End If
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'light green
            ElseIf Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(196, 215, 155) Then
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Value = True
                End If
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'bright green
            ElseIf Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(0, 176, 80) Then
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Value = True
                End If
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'orange
            ElseIf Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(255, 192, 0) Then
                Me.Controls("ToggleButton" & "2R" & rowIP).Value = True
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Enabled = False
                End If
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'blue
            ElseIf Me.Controls("TextBox" & "R" & rowIP).BackColor = RGB(146, 205, 220) Then
                Me.Controls("ToggleButton" & "3R" & rowIP).Value = True
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Enabled = False
                End If
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            'random color (lock entire row)
            Else
                'rows with 1a button
                If (rowIP = 3) Or (rowIP = 11) Or (rowIP = 13) Or (rowIP = 15) Or (rowIP = 19) Or (rowIP = 21) Or (rowIP = 23) Or (rowIP = 33) Then
                    Me.Controls("ToggleButton" & "1aR" & rowIP).Enabled = False
                End If
                Me.Controls("ToggleButton" & "1R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "2R" & rowIP).Enabled = False
                Me.Controls("ToggleButton" & "3R" & rowIP).Enabled = False
                Me.Controls("TextBox" & "R" & rowIP).Locked = True
            End If
            
            
            
        End If
lineNextDateCell:
    Next dateCell

    'enable toggles subs
    boolTogsAllowed = True

End Sub

'Engine Set Tracking SN Color
Private Sub ESTButton_Click()
    
    'Send Color to SN Column
    If Me.ESTButton.Value = True Then
        snTitleCell.Interior.Color = RGB(255, 255, 0)
    End If
    If Me.ESTButton.Value = False Then
        snTitleCell.Interior.Color = RGB(255, 255, 255)
    End If
    
    'Disable/Enable SlowButton
    If Me.ESTButton.Value = True Then
        Me.SlowButton.Enabled = False
    End If
    If Me.ESTButton.Value = False Then
        Me.SlowButton.Enabled = True
    End If
    
End Sub

'Slow Moving SN Color
Private Sub SlowButton_Click()
    
    'Send Color to SN Column
    If Me.SlowButton.Value = True Then
        snTitleCell.Interior.Color = RGB(244, 158, 228)
    End If
    If Me.SlowButton.Value = False Then
        snTitleCell.Interior.Color = RGB(255, 255, 255)
    End If
    
    'Disable/Enable ESTButton
    If Me.SlowButton.Value = True Then
        Me.ESTButton.Enabled = False
    End If
    If Me.SlowButton.Value = False Then
        Me.ESTButton.Enabled = True
    End If
    
End Sub

Private Sub MultiPage1_Change()

    'if changed to page 1
    If Me.MultiPage1.Value = 0 Then
        Application.Goto Worksheets("NEO 5322121").Cells(7, snSearchCol), Scroll:=True
    
    'if changed to page 2
    ElseIf Me.MultiPage1.Value = 1 Then
        Application.Goto Worksheets("NEO 5322121").Cells(23, snSearchCol), Scroll:=True
    End If

End Sub

'============================================Last Op Boxes================================================'
Private Sub LastOpTxtBox_Enter()

    Application.Goto Worksheets("NEO 5322121").Cells(51, snSearchCol)

End Sub

Private Sub LastOpTxtBox_Change()

    Worksheets("NEO 5322121").Cells(51, snSearchCol).Value = Me.LastOpTxtBox.Value

End Sub

Private Sub LastOpTxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Call MultiPage1_Change

End Sub

Private Sub LastDateTxtBox_Enter()

    Application.Goto Worksheets("NEO 5322121").Cells(52, snSearchCol)

End Sub

Private Sub LastDateTxtBox_Change()

    Worksheets("NEO 5322121").Cells(52, snSearchCol).Value = Me.LastDateTxtBox.Value

End Sub

Private Sub LastDateTxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Call MultiPage1_Change

End Sub

Private Sub LocationTxtBox_Enter()

    Application.Goto Worksheets("NEO 5322121").Cells(53, snSearchCol)

End Sub

Private Sub LocationTxtBox_Change()

    Worksheets("NEO 5322121").Cells(53, snSearchCol).Value = Me.LocationTxtBox.Value

End Sub

Private Sub LocationTxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Call MultiPage1_Change

End Sub
'========================================================================================================='

'=============================Green Toggles=============================='
Private Sub ToggleButton1R1_Click()
    boolTog = True: SNIPRow = 1: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R2_Click()
    boolTog = True: SNIPRow = 2: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR3_Click()
    boolTog = True: SNIPRow = 3: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R3_Click()
    boolTog = True: SNIPRow = 3: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R4_Click()
    boolTog = True: SNIPRow = 4: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R5_Click()
    boolTog = True: SNIPRow = 5: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R6_Click()
    boolTog = True: SNIPRow = 6: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R7_Click()
    boolTog = True: SNIPRow = 7: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R8_Click()
    boolTog = True: SNIPRow = 8: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R9_Click()
    boolTog = True: SNIPRow = 9: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R10_Click()
    boolTog = True: SNIPRow = 10: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR11_Click()
    boolTog = True: SNIPRow = 11: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R11_Click()
    boolTog = True: SNIPRow = 11: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R12_Click()
    boolTog = True: SNIPRow = 12: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR13_Click()
    boolTog = True: SNIPRow = 13: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R13_Click()
    boolTog = True: SNIPRow = 13: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R14_Click()
    boolTog = True: SNIPRow = 14: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR15_Click()
    boolTog = True: SNIPRow = 15: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R15_Click()
    boolTog = True: SNIPRow = 15: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R16_Click()
    boolTog = True: SNIPRow = 16: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R17_Click()
    boolTog = True: SNIPRow = 17: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R18_Click()
    boolTog = True: SNIPRow = 18: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR19_Click()
    boolTog = True: SNIPRow = 19: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R19_Click()
    boolTog = True: SNIPRow = 19: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R20_Click()
    boolTog = True: SNIPRow = 20: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR21_Click()
    boolTog = True: SNIPRow = 21: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R21_Click()
    boolTog = True: SNIPRow = 21: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R22_Click()
    boolTog = True: SNIPRow = 22: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR23_Click()
    boolTog = True: SNIPRow = 23: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R23_Click()
    boolTog = True: SNIPRow = 23: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R24_Click()
    boolTog = True: SNIPRow = 24: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R25_Click()
    boolTog = True: SNIPRow = 25: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R26_Click()
    boolTog = True: SNIPRow = 26: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R27_Click()
    boolTog = True: SNIPRow = 27: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R28_Click()
    boolTog = True: SNIPRow = 28: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R29_Click()
    boolTog = True: SNIPRow = 29: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R30_Click()
    boolTog = True: SNIPRow = 30: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R31_Click()
    boolTog = True: SNIPRow = 31: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R32_Click()
    boolTog = True: SNIPRow = 32: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1aR33_Click()
    boolTog = True: SNIPRow = 33: SNIPTog = "1a"
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton1R33_Click()
    boolTog = True: SNIPRow = 33: SNIPTog = "1"
    Call SNIPToggleButtons
End Sub
'==========================================================='

'=============================Orange Toggles=============================='
Private Sub ToggleButton2R1_Click()
    boolTog = True: SNIPRow = 1: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R2_Click()
    boolTog = True: SNIPRow = 2: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R3_Click()
    boolTog = True: SNIPRow = 3: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R4_Click()
    boolTog = True: SNIPRow = 4: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R5_Click()
    boolTog = True: SNIPRow = 5: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R6_Click()
    boolTog = True: SNIPRow = 6: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R7_Click()
    boolTog = True: SNIPRow = 7: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R8_Click()
    boolTog = True: SNIPRow = 8: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R9_Click()
    boolTog = True: SNIPRow = 9: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R10_Click()
    boolTog = True: SNIPRow = 10: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R11_Click()
    boolTog = True: SNIPRow = 11: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R12_Click()
    boolTog = True: SNIPRow = 12: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R13_Click()
    boolTog = True: SNIPRow = 13: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R14_Click()
    boolTog = True: SNIPRow = 14: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R15_Click()
    boolTog = True: SNIPRow = 15: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R16_Click()
    boolTog = True: SNIPRow = 16: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R17_Click()
    boolTog = True: SNIPRow = 17: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R18_Click()
    boolTog = True: SNIPRow = 18: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R19_Click()
    boolTog = True: SNIPRow = 19: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R20_Click()
    boolTog = True: SNIPRow = 20: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R21_Click()
    boolTog = True: SNIPRow = 21: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R22_Click()
    boolTog = True: SNIPRow = 22: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R23_Click()
    boolTog = True: SNIPRow = 23: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R24_Click()
    boolTog = True: SNIPRow = 24: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R25_Click()
    boolTog = True: SNIPRow = 25: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R26_Click()
    boolTog = True: SNIPRow = 26: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R27_Click()
    boolTog = True: SNIPRow = 27: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R28_Click()
    boolTog = True: SNIPRow = 28: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R29_Click()
    boolTog = True: SNIPRow = 29: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R30_Click()
    boolTog = True: SNIPRow = 30: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R31_Click()
    boolTog = True: SNIPRow = 31: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R32_Click()
    boolTog = True: SNIPRow = 32: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton2R33_Click()
    boolTog = True: SNIPRow = 33: SNIPTog = "2"
    'call togglebuttons sub
    Call SNIPToggleButtons
End Sub
'==========================================================='

'=============================Blue Toggles=============================='
Private Sub ToggleButton3R1_Click()
    boolTog = True: SNIPRow = 1: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R2_Click()
    boolTog = True: SNIPRow = 2: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R3_Click()
    boolTog = True: SNIPRow = 3: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R4_Click()
    boolTog = True: SNIPRow = 4: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R5_Click()
    boolTog = True: SNIPRow = 5: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R6_Click()
    boolTog = True: SNIPRow = 6: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R7_Click()
    boolTog = True: SNIPRow = 7: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R8_Click()
    boolTog = True: SNIPRow = 8: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R9_Click()
    boolTog = True: SNIPRow = 9: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R10_Click()
    boolTog = True: SNIPRow = 10: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R11_Click()
    boolTog = True: SNIPRow = 11: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R12_Click()
    boolTog = True: SNIPRow = 12: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R13_Click()
    boolTog = True: SNIPRow = 13: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R14_Click()
    boolTog = True: SNIPRow = 14: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R15_Click()
    boolTog = True: SNIPRow = 15: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R16_Click()
    boolTog = True: SNIPRow = 16: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R17_Click()
    boolTog = True: SNIPRow = 17: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R18_Click()
    boolTog = True: SNIPRow = 18: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R19_Click()
    boolTog = True: SNIPRow = 19: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R20_Click()
    boolTog = True: SNIPRow = 20: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R21_Click()
    boolTog = True: SNIPRow = 21: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R22_Click()
    boolTog = True: SNIPRow = 22: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R23_Click()
    boolTog = True: SNIPRow = 23: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R24_Click()
    boolTog = True: SNIPRow = 24: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R25_Click()
    boolTog = True: SNIPRow = 25: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R26_Click()
    boolTog = True: SNIPRow = 26: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R27_Click()
    boolTog = True: SNIPRow = 27: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R28_Click()
    boolTog = True: SNIPRow = 28: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R29_Click()
    boolTog = True: SNIPRow = 29: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R30_Click()
    boolTog = True: SNIPRow = 30: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R31_Click()
    boolTog = True: SNIPRow = 31: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R32_Click()
    boolTog = True: SNIPRow = 32: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub

Private Sub ToggleButton3R33_Click()
    boolTog = True: SNIPRow = 33: SNIPTog = "3": boolRTOTog = True
    Call SNIPToggleButtons
End Sub
'==========================================================='

Private Sub SNIPToggleButtons()

    'For disable rows beneath section
    Dim i As Double
    Dim SNIPSearchCurrentRowTrans As Double
    'define SNIPSearchCurrentRowTrans
    If (SNIPSearchCurrentRow >= 7) And (SNIPSearchCurrentRow <= 33) Then
        SNIPSearchCurrentRowTrans = SNIPSearchCurrentRow - 6
    ElseIf (SNIPSearchCurrentRow >= 38) And (SNIPSearchCurrentRow <= 44) Then
        SNIPSearchCurrentRowTrans = SNIPSearchCurrentRow - 10
    End If

    'Toggle Buttons
        'if toggles are allowed
        If boolTogsAllowed = True Then
            'Button is toggled
                If (Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).Enabled = True) And (Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).Value = True) Then
                    
                    'textbox lock
                    Me.Controls("TextBox" & "R" & SNIPRow).Locked = True
                    
                    'Disable Buttons
                        'rows with 1a button
                    If (SNIPRow = 3) Or (SNIPRow = 11) Or (SNIPRow = 13) Or (SNIPRow = 15) Or (SNIPRow = 19) Or (SNIPRow = 21) Or (SNIPRow = 23) Or (SNIPRow = 33) Then
                        Me.Controls("ToggleButton" & "1aR" & SNIPRow).Enabled = False
                    End If
                    Me.Controls("ToggleButton" & 1 & "R" & SNIPRow).Enabled = False
                    Me.Controls("ToggleButton" & 2 & "R" & SNIPRow).Enabled = False
                    Me.Controls("ToggleButton" & 3 & "R" & SNIPRow).Enabled = False
                    Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).Enabled = True
                    
                    'set variables
                    txtBoxTxt = Me.Controls("TextBox" & "R" & SNIPRow).Value
                    txtBoxClr = Me.Controls("TextBox" & "R" & SNIPRow).BackColor
                    togClr = Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).BackColor
                    
                    'call textbox color sub
                    If togClr = RGB(146, 208, 80) Then
                        Call SNIPTxtBoxGreen
                    ElseIf togClr = RGB(79, 98, 40) Then
                        Call SNIPTxtBoxGreenDark
                    ElseIf togClr = RGB(196, 215, 155) Then
                        Call SNIPTxtBoxGreenLight
                    ElseIf togClr = RGB(0, 176, 80) Then
                        Call SNIPTxtBoxGreenBright
                    ElseIf togClr = RGB(255, 192, 0) Then
                        Call SNIPTxtBoxOrange
                    ElseIf togClr = RGB(146, 205, 220) Then
                        Call SNIPTxtBoxBlue
                    End If
                    
                    'finalize variables
                    Me.Controls("TextBox" & "R" & SNIPRow).Value = txtBoxTxt
                    Me.Controls("TextBox" & "R" & SNIPRow).BackColor = txtBoxClr
                    
                    'Disable buttons in all rows beneath that are not already locked
                    For i = (SNIPRow + 1) To (SNIPSearchCurrentRowTrans - 1)
                        'if row is not locked already
                        If Me.Controls("TextBox" & "R" & i).Locked = False Then
                            
                            'lock and color textbox
                            Call SNIPTxtBoxGreen
                            Me.Controls("TextBox" & "R" & i).Locked = True
                            Me.Controls("TextBox" & "R" & i).BackColor = txtBoxClr
                            'fill corresponding tracker cell
                            If i <= 27 Then
                                If Me.Controls("TextBox" & "R" & i).Value = "" Then
                                    Worksheets("NEO 5322121").Cells((i + 6), snTitleCell.Column).Value = ""
                                End If
                                Worksheets("NEO 5322121").Cells((i + 6), snTitleCell.Column).Interior.Color = txtBoxClr
                            ElseIf (i > 27) And (i <= 33) Then
                                If Me.Controls("TextBox" & "R" & i).Value = "" Then
                                    Worksheets("NEO 5322121").Cells((i + 10), snTitleCell.Column).Value = ""
                                End If
                                Worksheets("NEO 5322121").Cells((i + 10), snTitleCell.Column).Interior.Color = txtBoxClr
                            End If
                            
                            'if txtbox not empty, toggle corresponding colored button
                            If Not (Me.Controls("TextBox" & "R" & i).Value = "") Then
                                'disable toggling
                                boolTogsAllowed = False
                                'Green
                                If Me.Controls("TextBox" & "R" & i).BackColor = RGB(146, 208, 80) Then
                                    Me.Controls("ToggleButton" & 1 & "R" & i).Value = True
                                'Dark Green
                                ElseIf Me.Controls("TextBox" & "R" & i).BackColor = RGB(79, 98, 40) Then
                                    Me.Controls("ToggleButton" & 1 & "aR" & i).Value = True
                                'Light Green
                                ElseIf Me.Controls("TextBox" & "R" & i).BackColor = RGB(196, 215, 155) Then
                                    Me.Controls("ToggleButton" & 1 & "aR" & i).Value = True
                                'Bright Green
                                ElseIf Me.Controls("TextBox" & "R" & i).BackColor = RGB(0, 176, 80) Then
                                    Me.Controls("ToggleButton" & 1 & "aR" & i).Value = True
                                'Orange
                                ElseIf Me.Controls("TextBox" & "R" & i).BackColor = RGB(255, 192, 0) Then
                                    Me.Controls("ToggleButton" & 2 & "R" & i).Value = True
                                'Blue
                                ElseIf Me.Controls("TextBox" & "R" & i).BackColor = RGB(146, 205, 220) Then
                                    Me.Controls("ToggleButton" & 3 & "R" & i).Value = True
                                End If
                                'enable toggling
                                boolTogsAllowed = True
                            End If
                            
                            'disable buttons
                                'rows with 1a button
                            If (i = 3) Or (i = 11) Or (i = 13) Or (i = 15) Or (i = 19) Or (i = 21) Or (i = 23) Or (i = 33) Then
                                If Me.Controls("ToggleButton" & 1 & "aR" & i).Value = False Then
                                    Me.Controls("ToggleButton" & 1 & "aR" & i).Enabled = False
                                End If
                            End If
                            If Me.Controls("ToggleButton" & 1 & "R" & i).Value = False Then
                                Me.Controls("ToggleButton" & 1 & "R" & i).Enabled = False
                            End If
                            If Me.Controls("ToggleButton" & 2 & "R" & i).Value = False Then
                                Me.Controls("ToggleButton" & 2 & "R" & i).Enabled = False
                            End If
                            If Me.Controls("ToggleButton" & 3 & "R" & i).Value = False Then
                                Me.Controls("ToggleButton" & 3 & "R" & i).Enabled = False
                            End If
                            
                        End If
                    Next i
                    
                    
                'Button is untoggled
                ElseIf (Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).Enabled = True) And (Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).Value = False) Then
                    
                    'textbox unlock
                    Me.Controls("TextBox" & "R" & SNIPRow).Locked = False
                    
                    'enable buttons
                        'rows with 1a button
                    If (SNIPRow = 3) Or (SNIPRow = 11) Or (SNIPRow = 13) Or (SNIPRow = 15) Or (SNIPRow = 19) Or (SNIPRow = 21) Or (SNIPRow = 23) Or (SNIPRow = 33) Then
                        Me.Controls("ToggleButton" & "1aR" & SNIPRow).Enabled = True
                    End If
                    Me.Controls("ToggleButton" & 1 & "R" & SNIPRow).Enabled = True
                    Me.Controls("ToggleButton" & 2 & "R" & SNIPRow).Enabled = True
                    Me.Controls("ToggleButton" & 3 & "R" & SNIPRow).Enabled = True
                    
                    'set variables
                    boolTog = False
                    txtBoxTxt = Me.Controls("TextBox" & "R" & SNIPRow).Value
                    txtBoxClr = Me.Controls("TextBox" & "R" & SNIPRow).BackColor
                    togClr = Me.Controls("ToggleButton" & SNIPTog & "R" & SNIPRow).BackColor
                    
                    'call textbox color sub
                    If togClr = RGB(146, 208, 80) Then
                        Call SNIPTxtBoxGreen
                    ElseIf togClr = RGB(79, 98, 40) Then
                        Call SNIPTxtBoxGreenDark
                    ElseIf togClr = RGB(196, 215, 155) Then
                        Call SNIPTxtBoxGreenLight
                    ElseIf togClr = RGB(0, 176, 80) Then
                        Call SNIPTxtBoxGreenBright
                    ElseIf togClr = RGB(255, 192, 0) Then
                        Call SNIPTxtBoxOrange
                    ElseIf togClr = RGB(146, 205, 220) Then
                        Call SNIPTxtBoxBlue
                    End If
                    
                    'finalize variables
                    Me.Controls("TextBox" & "R" & SNIPRow).Value = txtBoxTxt
                    Me.Controls("TextBox" & "R" & SNIPRow).BackColor = txtBoxClr
                    
                    'Enable buttons in all rows beneath that were not already locked
                    For i = (SNIPRow + 1) To (SNIPSearchCurrentRowTrans - 1)
                        'if row was newly locked
                        If (Me.Controls("TextBox" & "R" & i).Locked = True) And (Me.Controls("ToggleButton" & 1 & "aR" & i).Value = False) And (Me.Controls("ToggleButton" & 1 & "R" & i).Value = False) And (Me.Controls("ToggleButton" & 2 & "R" & i).Value = False) And (Me.Controls("ToggleButton" & 3 & "R" & i).Value = False) Then
                            'unlock and uncolor textbox
                            boolTog = False
                            Call SNIPTxtBoxGreen
                            Me.Controls("TextBox" & "R" & i).Locked = False
                            Me.Controls("TextBox" & "R" & i).BackColor = txtBoxClr
                            'remove fill of corresponding tracker cell
                            If i <= 27 Then
                                Worksheets("NEO 5322121").Cells((i + 6), snTitleCell.Column).Interior.Color = txtBoxClr
                            ElseIf (i > 27) And (i <= 33) Then
                                Worksheets("NEO 5322121").Cells((i + 10), snTitleCell.Column).Interior.Color = txtBoxClr
                            End If
                            
                            'enable buttons
                                'rows with 1a button
                            If (i = 3) Or (i = 11) Or (i = 13) Or (i = 15) Or (i = 19) Or (i = 21) Or (i = 23) Or (i = 33) Then
                                Me.Controls("ToggleButton" & 1 & "aR" & i).Enabled = True
                            End If
                            Me.Controls("ToggleButton" & 1 & "R" & i).Enabled = True
                            Me.Controls("ToggleButton" & 2 & "R" & i).Enabled = True
                            Me.Controls("ToggleButton" & 3 & "R" & i).Enabled = True
                        End If
                    Next i
                    
                End If
                
                
                
                'Update Serial Number Column
                
                Dim TrackerRow As Double
                
                'for SNIP rows 1 through 27
                If (SNIPRow >= 1) And (SNIPRow <= 27) Then
                    TrackerRow = SNIPRow + 6
                'for SNIP rows 28 through 33
                ElseIf (SNIPRow >= 28) And (SNIPRow <= 33) Then
                    TrackerRow = SNIPRow + 10
                End If
                
                'write to SNcol
                Worksheets("NEO 5322121").Cells(TrackerRow, snSearchCol).Value = Me.Controls("TextBox" & "R" & SNIPRow).Value
                Worksheets("NEO 5322121").Cells(TrackerRow, snSearchCol).Interior.Color = Me.Controls("TextBox" & "R" & SNIPRow).BackColor
                    
                    'if rto button hit, search info page RTO buttons for topmost toggled button
                    If boolRTOTog Then
                        Dim intTopRTO As Double
                        For intTopRTO = 1 To 33
                            'if a toggled rto button is found
                            If Me.Controls("ToggleButton" & 3 & "R" & intTopRTO).Value = True Then
                                Worksheets("NEO 5322121").Cells(50, snSearchCol).Value = Me.Controls("TextBox" & "R" & intTopRTO).Value
                                Worksheets("NEO 5322121").Cells(50, snSearchCol).Interior.Color = Me.Controls("TextBox" & "R" & intTopRTO).BackColor
                                Worksheets("NEO 5322121").Cells(49, snSearchCol).Value = "R2O"
                                Worksheets("NEO 5322121").Cells(49, snSearchCol).Interior.Color = Me.Controls("TextBox" & "R" & intTopRTO).BackColor
                                Exit For
                            'if no rto buttons are toggled
                            Else
                                Worksheets("NEO 5322121").Cells(50, snSearchCol).Value = ""
                                Worksheets("NEO 5322121").Cells(50, snSearchCol).Interior.Color = clrBlank
                                Worksheets("NEO 5322121").Cells(49, snSearchCol).Value = ""
                                Worksheets("NEO 5322121").Cells(49, snSearchCol).Interior.Color = clrBlank
                            End If
                        Next intTopRTO
                        'reset boolRTOTog
                        boolRTOTog = False
                    End If
        
        End If

    'reset boolTog
    boolTog = False

End Sub

Private Sub SearchNewButton_Click()

    '============================= Check for incorrect textbox values ================================='
    Dim t As Double
    Dim s As Double
    
    'iterate through date textboxes
    For t = 1 To 33
        'iterate through textbox value
        For s = 1 To Len(Me.Controls("TextBox" & "R" & t).Value)
            'if incorrect value found
            If (InStr(1, "0123456789/", Mid(Me.Controls("TextBox" & "R" & t).Value, s, 1)) = 0) Or (InStr(1, "0123456789", Right(Me.Controls("TextBox" & "R" & t).Value, 1)) = 0) Then
                'user error message
                MsgBox "You have entered an incorect character in Text Box #" & t & ". Please revise before continuing."
                Exit Sub
            End If
        Next s
    Next t
    
    'check last date seen text box
    For s = 1 To Len(LastDateTxtBox.Value)
        'if incorrect value found
            If (InStr(1, "0123456789/", Mid(LastDateTxtBox.Value, s, 1)) = 0) Or (InStr(1, "0123456789", Right(LastDateTxtBox.Value, 1)) = 0) Then
                'user error message
                MsgBox "You have entered an incorect character in the Last Date Seen Text Box. Please revise before continuing."
                Exit Sub
            End If
    Next s
    '============================= Check for incorrect textbox values ================================='

    'load QCPartInfo form
    Set frmSNQCPartInfo = New SN_QCPartInfo
    frmSNQCPartInfo.ButtonConfirm.SetFocus
    
    'look for orange cell
    Dim badcll As Range
    For Each badcll In Worksheets("NEO 5322121").Range(Cells(7, snTitleCell.Column), Cells(43, snTitleCell.Column))
        If (badcll.Interior.Color = RGB(255, 192, 0)) And ((badcll.Row <= 33) Or (badcll.Row >= 38)) Then
            'default QC Status to orange bad if bad cell found
            Worksheets("NEO 5322121").Cells(54, snSearchCol).Value = "Bad"
            Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color = RGB(255, 192, 0)
            frmSNQCPartInfo.ToggleButtonBad.Value = True
            Exit For
        End If
    Next badcll
    
    'initialize textboxes and togglebutton
    frmSNQCPartInfo.TextBoxQCStatus.Value = Worksheets("NEO 5322121").Cells(54, snSearchCol).Value
    frmSNQCPartInfo.TextBoxQCStatus.BackColor = Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color
    frmSNQCPartInfo.TextBoxRiskProfile.Value = Worksheets("NEO 5322121").Cells(56, snSearchCol).Value
    frmSNQCPartInfo.TextBoxRiskProfile.BackColor = Worksheets("NEO 5322121").Cells(56, snSearchCol).Interior.Color
    frmSNQCPartInfo.ToggleButtonBad.BackColor = RGB(255, 192, 0)
    'show form
    frmSNQCPartInfo.Show
    
    '========================WATERFALL ARRAY==============================='
    'Copy column to updated array
    Dim w As Range
    For Each w In Worksheets("NEO 5322121").Range(Cells(1, snTitleCell.Column), Cells(56, snTitleCell.Column))
        '==========================='
        'if error cell flagged
        If arraySNBackupVal(w.Row) = "!!!FLAGGED AS ERROR!!!" Then: GoTo lineErrorFoundw
        '==========================='
        'assign value to array
        arrayWaterfallVal(w.Row) = w.Value
        'encountering error
        If False Then
lineErrorFoundw:
            arrayWaterfallVal(w.Row) = "!!!FLAGGED AS ERROR!!!"
            'reset booleans
            boolFlagCell = False
            boolFlagCellClear = False
        End If
        arrayWaterfallClr(w.Row) = w.Interior.Color
    Next w
    
    'SN Waterfall Cut Column
    arrayWaterfallVal(57) = snTitleCell.Column
    
    '======================================================================'
    
    'hide form
    Me.Hide
    boolCanceled = True
    
    'Waterfall Tracker
    Call WaterFallSN
    
    'call snsearch
    If intSNMLType <> 5 Then
        Call SNSearchBox
    'call snQCtoWIP
    ElseIf intSNMLType = 5 Then
        Application.Goto Worksheets("Quality Clinic").Range("C6")
        Call SNQCtoWIP
    End If

End Sub

Private Sub UpdateGroupButton_Click()

    '============================= Check for incorrect textbox values ================================='
    Dim t As Double
    Dim s As Double
    
    'iterate through date textboxes
    For t = 1 To 33
        'iterate through textbox value
        For s = 1 To Len(Me.Controls("TextBox" & "R" & t).Value)
            'if incorrect value found
            If (InStr(1, "0123456789/", Mid(Me.Controls("TextBox" & "R" & t).Value, s, 1)) = 0) Or (InStr(1, "0123456789", Right(Me.Controls("TextBox" & "R" & t).Value, 1)) = 0) Then
                'user error message
                MsgBox "You have entered an incorect character in Text Box #" & t & ". Please revise before continuing."
                Exit Sub
            End If
        Next s
    Next t
    
    'check last date seen text box
    For s = 1 To Len(LastDateTxtBox.Value)
        'if incorrect value found
            If (InStr(1, "0123456789/", Mid(LastDateTxtBox.Value, s, 1)) = 0) Or (InStr(1, "0123456789", Right(LastDateTxtBox.Value, 1)) = 0) Then
                'user error message
                MsgBox "You have entered an incorect character in the Last Date Seen Text Box. Please revise before continuing."
                Exit Sub
            End If
    Next s
    '============================= Check for incorrect textbox values ================================='

    'load QCPartInfo form
    Set frmSNQCPartInfo = New SN_QCPartInfo
    frmSNQCPartInfo.ButtonConfirm.SetFocus
    
    'look for orange cell
    Dim badcll As Range
    For Each badcll In Worksheets("NEO 5322121").Range(Cells(7, snTitleCell.Column), Cells(43, snTitleCell.Column))
        If (badcll.Interior.Color = RGB(255, 192, 0)) And ((badcll.Row <= 33) Or (badcll.Row >= 38)) Then
            'default QC Status to orange bad if bad cell found
            Worksheets("NEO 5322121").Cells(54, snSearchCol).Value = "Bad"
            Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color = RGB(255, 192, 0)
            frmSNQCPartInfo.ToggleButtonBad.Value = True
            Exit For
        End If
    Next badcll
    
    'initialize textboxes and togglebutton
    frmSNQCPartInfo.TextBoxQCStatus.Value = Worksheets("NEO 5322121").Cells(54, snSearchCol).Value
    frmSNQCPartInfo.TextBoxQCStatus.BackColor = Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color
    frmSNQCPartInfo.TextBoxRiskProfile.Value = Worksheets("NEO 5322121").Cells(56, snSearchCol).Value
    frmSNQCPartInfo.TextBoxRiskProfile.BackColor = Worksheets("NEO 5322121").Cells(56, snSearchCol).Interior.Color
    frmSNQCPartInfo.ToggleButtonBad.BackColor = RGB(255, 192, 0)
    If Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color = RGB(255, 192, 0) Then
        frmSNQCPartInfo.ToggleButtonBad.Value = True
    End If
    'show form
    frmSNQCPartInfo.Show

    '===================UPDATED SN AND WATERFALL ARRAY====================='
    'Copy column to updated array
    Dim n As Range
    For Each n In Worksheets("NEO 5322121").Range(Cells(1, snTitleCell.Column), Cells(56, snTitleCell.Column))
        '==========================='
        'if error cell flagged
        If arraySNBackupVal(n.Row) = "!!!FLAGGED AS ERROR!!!" Then: GoTo lineErrorFoundn
        '==========================='
        'assign value to array
        arraySNUpdatedVal(n.Row) = n.Value
        'encountering error
        If False Then
lineErrorFoundn:
            arraySNUpdatedVal(n.Row) = "!!!FLAGGED AS ERROR!!!"
        End If
        arrayWaterfallVal(n.Row) = arraySNUpdatedVal(n.Row)
        arraySNUpdatedClr(n.Row) = n.Interior.Color
        arrayWaterfallClr(n.Row) = n.Interior.Color
    Next n
    
    'SN Waterfall Cut Column
    arrayWaterfallVal(57) = snTitleCell.Column
    
    '======================================================================'
    '========================UPDATES ONLY ARRAY============================'
    'find differences between SN backup and updated
    Dim d As Double
    For d = 1 To 56
        'if value difference found
        If Not (arraySNUpdatedVal(d) = arraySNBackupVal(d)) Then
            'prevent copying of Serial Number
            If d <> 6 Then: arrayUpdatesOnlyVal(d) = arraySNUpdatedVal(d)
            arrayUpdatesOnlyClr(d) = arraySNUpdatedClr(d)
        End If
        'if color difference found
        If Not (arraySNUpdatedClr(d) = arraySNBackupClr(d)) Then
            'prevent copying of Serial Number
            If d <> 6 Then: arrayUpdatesOnlyVal(d) = arraySNUpdatedVal(d)
            arrayUpdatesOnlyClr(d) = arraySNUpdatedClr(d)
        End If
    Next d
    '======================================================================'

    'hide info page
    Me.Hide
    
    'Waterfall Tracker
    Call WaterFallSN
    
    'call snduplicate
    boolSNDTGFirstRun = True
    Call SNDuplicateToGroup

End Sub

Private Sub CancelButton_Click()

    '======================================================================'
    'Copy backup array to SN column
    Dim s As Range
    For Each s In Worksheets("NEO 5322121").Range(Cells(1, snTitleCell.Column), Cells(56, snTitleCell.Column))
        'ignore flagged as error entry
        If Not (arraySNBackupVal(s.Row) = "!!!FLAGGED AS ERROR!!!") Then
            s.Value = arraySNBackupVal(s.Row)
            s.Interior.Color = arraySNBackupClr(s.Row)
        End If
    Next s
    '======================================================================'
    
    'hide form
    Me.Hide
    boolCanceled = True
    
    'call snsearch
    If intSNMLType <> 5 Then
        Call SNSearchBox
    'call snQCtoWIP
    ElseIf intSNMLType = 5 Then
        Application.Goto Worksheets("Quality Clinic").Range("C6")
        Call SNQCtoWIP
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If

End Sub
