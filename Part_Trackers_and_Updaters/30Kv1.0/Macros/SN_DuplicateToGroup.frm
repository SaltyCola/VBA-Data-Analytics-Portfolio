VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_DuplicateToGroup 
   Caption         =   "Duplicate Most Recent Changes"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3360
   OleObjectBlob   =   "SN_DuplicateToGroup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_DuplicateToGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConfirmButton_Click()

    'find SN from listbox in tracker
    Dim lstInt As Double
    For lstInt = 0 To (Me.ListBox1.ListCount - 1)
        
        '==============================Search Tracker==============================='
        Dim trgtStr As String
        Dim DTGsnRng As Range
        Dim snCell As Range
        Dim snCol As Long
        Dim snRow As Long
        Dim snTxtRaw As String
        Dim snTxt As String
        Dim snTxtPrefix
        Dim snTxtU As String
        Dim snTxtL As String
    
        'Define snRng
        If Me.ConfirmButton.Caption = "Waterfall QC" Then
            Set DTGsnRng = Worksheets("Quality Clinic").Range("6:6")
        Else
            Set DTGsnRng = Worksheets("NEO 5322121").Range("6:6")
        End If
    
        'Define variables
        trgtStr = Me.ListBox1.List(lstInt)
        
        'serial number correct format
            'search sn row in tracker
            For Each snCell In DTGsnRng
                snTxtRaw = snCell.Value
                snRow = snCell.Row
                snCol = snCell.Column
                'only look at SN's
                If Len(snTxtRaw) > 5 Then
                    snTxt = Right(snTxtRaw, (Len(snTxtRaw) - 5))
                    snTxtPrefix = Mid(snTxt, 1, 1)
                    snTxtU = UCase(snTxtPrefix) & Mid(snTxt, 2)
                    snTxtL = LCase(snTxtPrefix) & Mid(snTxt, 2)
                    
                    'On serial number found:
                    If (Right(trgtStr, 5) = snTxtU) Or (Right(trgtStr, 5) = snTxtL) Then
                        
                        'make sn ListedSNCell
                        Set SNDTGListedSNCell = snCell
                        GoTo lineUpdateGroup
                    
                    End If
                End If
            Next snCell
        '==========================================================================='
lineUpdateGroup:
        '======================================================================'
        'Copy updated array to SN column
        Dim u As Range
        Dim uRng As Range
        
        'set uRng
        If Me.ConfirmButton.Caption = "Waterfall QC" Then
            Set uRng = Worksheets("Quality Clinic").Range(Cells(1, SNDTGListedSNCell.Column), Cells(56, SNDTGListedSNCell.Column))
        Else
            Set uRng = Worksheets("NEO 5322121").Range(Cells(1, SNDTGListedSNCell.Column), Cells(56, SNDTGListedSNCell.Column))
        End If
        
        For Each u In uRng
                'coming from WIP updater
            'ignore flagged as error entry
            If Not (boolWtrFllDTG) Then
                '================================================================================='
                'initialize boolFlagCell and boolFlagCellClear
                boolFlagCell = False
                boolFlagCellClear = False
                'search for errors
                If (IsError(u.Value)) Then
                    'Load SN_DeleteError form
                    Set frmSNDeleteErrorDTG = New SN_DeleteError
                    boolCanceled = False
                    'initialize txtboxes
                    frmSNDeleteErrorDTG.TextBox1.Value = u.Address(False, False)
                    'define error cell range and go to error cell
                    rngErrorCell = u.Address
                    If Me.ConfirmButton.Caption = "Waterfall QC" Then
                        Application.Goto Worksheets("Quality Clinic").Range(rngErrorCell)
                    Else
                        Application.Goto Worksheets("NEO 5322121").Range(rngErrorCell)
                    End If
                    'show form
                    frmSNDeleteErrorDTG.Show
                    'if SN_DeleteError is exitted
                    If boolCanceled Then
                        Exit Sub
                    End If
                End If
                'if error cell flagged
                If boolFlagCell Then: GoTo lineDTGErrorFound
                '================================================================================='
                'if change in value
                If Not (arrayUpdatesOnlyVal(u.Row) = "!===N/A===!") And Not (arrayUpdatesOnlyVal(u.Row) = "!!!FLAGGED AS ERROR!!!") Then
                    u.Value = arrayUpdatesOnlyVal(u.Row)
                    arrayWaterfallVal(u.Row) = arrayUpdatesOnlyVal(u.Row)
                'if no change in value
                ElseIf arrayUpdatesOnlyVal(u.Row) = "!===N/A===!" Then
                    arrayWaterfallVal(u.Row) = u.Value
                'if error in value
                ElseIf False Then
lineDTGErrorFound:
                    arrayWaterfallVal(u.Row) = "!!!FLAGGED AS ERROR!!!"
                    'reset booleans
                    boolFlagCell = False
                    boolFlagCellClear = False
                End If
                'if change in color
                If Not (arrayUpdatesOnlyClr(u.Row) = 102030405) Then
                    u.Interior.Color = arrayUpdatesOnlyClr(u.Row)
                    arrayWaterfallClr(u.Row) = arrayUpdatesOnlyClr(u.Row)
                'if no change in color
                ElseIf arrayUpdatesOnlyClr(u.Row) = 102030405 Then
                    arrayWaterfallClr(u.Row) = u.Interior.Color
                End If
            
                'coming from main menu (Waterfalling Entire Tracker)
            ElseIf boolWtrFllDTG Then
                '================================================================================='
                'initialize boolFlagCell and boolFlagCellClear
                boolFlagCell = False
                boolFlagCellClear = False
                'search for errors
                If (IsError(u.Value)) Then
                    'Load SN_DeleteError form
                    Set frmSNDeleteErrorWtrFllTrkr = New SN_DeleteError
                    boolCanceled = False
                    'initialize txtboxes
                    frmSNDeleteErrorWtrFllTrkr.TextBox1.Value = u.Address(False, False)
                    'define error cell range and go to error cell
                    rngErrorCell = u.Address
                    If Me.ConfirmButton.Caption = "Waterfall QC" Then
                        Application.Goto Worksheets("Quality Clinic").Range(rngErrorCell)
                    Else
                        Application.Goto Worksheets("NEO 5322121").Range(rngErrorCell)
                    End If
                    'show form
                    frmSNDeleteErrorWtrFllTrkr.Show
                    'if SN_DeleteError is exitted
                    If boolCanceled Then
                        Exit Sub
                    End If
                End If
                'if error cell flagged
                If boolFlagCell Then: GoTo lineDTGErrorFoundwtrfll
                '================================================================================='
                arrayWaterfallVal(u.Row) = u.Value
                'encountering error
                If False Then
lineDTGErrorFoundwtrfll:
                    arrayWaterfallVal(u.Row) = "!!!FLAGGED AS ERROR!!!"
                    'reset booleans
                    boolFlagCell = False
                    boolFlagCellClear = False
                End If
                arrayWaterfallClr(u.Row) = u.Interior.Color
            End If
        Next u
    
        'SN Waterfall Cut Column
        arrayWaterfallVal(57) = SNDTGListedSNCell.Column
        
        '======================================================================'
        
        'reset booleans if waterfalling entire tracker
        If boolWtrFllDTG Then
            boolDoNotClose = True
        End If
        If lstInt = (Me.ListBox1.ListCount - 1) Then
            boolDoNotClose = False
        End If
        If lstInt = 0 Then
            boolWtrFllDTGFirstRun = True
        ElseIf lstInt > 0 Then
            boolWtrFllDTGFirstRun = False
        End If
        
        'hide form and call waterfaller
        Me.Hide
        If Me.ConfirmButton.Caption = "Waterfall QC" Then
            Call WaterFallQC
        Else
            Call WaterFallSN
        End If
        
        'Reset arraywaterfallval/clr
        ReDim arrayWaterfallVal(57)
        ReDim arrayWaterfallClr(56)

    Next lstInt

    'completed message
    MsgBox ("All Serial Numbers Updated!"), , "Update Complete"
    
    'call snsearchbox if from info page
    If Not boolWtrFllDTG Then
        Call SNSearchBox
    'call main menu if from waterfall entire tracker
    ElseIf boolWtrFllDTG Then
        boolWtrFllDTG = False
        Call TrackerMainMenu
    End If

End Sub

Private Sub TextBox_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)

    If keycode = vbKeyReturn Then
        
        Dim trgtStr As String
        Dim snRng As Range
        Dim sn As Range
        Dim snCol As Long
        Dim snRow As Long
        Dim snTxtRaw As String
        Dim snTxt As String
        Dim snTxtPrefix
        Dim snTxtU As String
        Dim snTxtL As String
    
        'Define snRng
        Set snRng = Worksheets("NEO 5322121").Range("6:6")
    
        'Define variables
        trgtStr = Me.TextBox.Value
        
        'Length / Format Error Msgs
        If (Len(trgtStr) < 4) Or (Len(trgtStr) > 5) Then
            Me.Hide
            MsgBox ("Please enter the correct serial number format. (i.e. J0101 or 0101)"), , "Length Error"
            Me.TextBox.Value = ""
            Me.Show
            Exit Sub
        ElseIf (Len(trgtStr) = 5) And (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(trgtStr, 1, 1))) = 0) Then
            Me.Hide
            MsgBox ("The first character of a five character entry must be a letter. (i.e. J0101)"), , "Format Error"
            Me.TextBox.Value = ""
            Me.Show
            Exit Sub
        ElseIf (InStr("0123456789", Mid(Right(trgtStr, 4), 1, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 2, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 3, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 4, 1)) = 0) Then
            Me.Hide
            MsgBox ("The final 4 characters must be numbers. (i.e. J0101 or 0101)"), , "Format Error"
            Me.TextBox.Value = ""
            Me.Show
            Exit Sub
        
        'serial number correct format
        Else
            'search sn row in tracker
            For Each sn In snRng
                snTxtRaw = sn.Value
                snRow = sn.Row
                snCol = sn.Column
                'only look at SN's
                If Len(snTxtRaw) > 5 Then
                    snTxt = Right(snTxtRaw, (Len(snTxtRaw) - 5))
                    snTxtPrefix = Mid(snTxt, 1, 1)
                    snTxtU = UCase(snTxtPrefix) & Mid(snTxt, 2)
                    snTxtL = LCase(snTxtPrefix) & Mid(snTxt, 2)
                    
                    'On serial number found:
                    If ((trgtStr = snTxtU) Or (trgtStr = Right(snTxtU, 4))) Or ((trgtStr = snTxtL) Or (trgtStr = Right(snTxtL, 4))) Then
                        'initialize snml type
                        intSNMLType = 2
                        'first sn match
                        If SNMLArrayCnt = 0 Then
                            'populate matchlist array
                            SNMLArray(UBound(SNMLArray)) = sn.Value
                            SNMLArrayCnt = SNMLArrayCnt + 1
                        'every match after the first
                        ElseIf SNMLArrayCnt > 0 Then
                            'redimension array
                            ReDim Preserve SNMLArray(SNMLArrayCnt)
                            'populate matchlist array
                            SNMLArray(UBound(SNMLArray)) = sn.Value
                            SNMLArrayCnt = SNMLArrayCnt + 1
                        End If
                    
                    End If
                End If
            Next sn
            
            'hide form
            Me.Hide
            
            'initialize snml type
            intSNMLType = 2
            
            'Call Match list
            Call SNMatchList
            
        End If
        
    'clear out textbox
    Me.TextBox.Value = ""
    End If

End Sub

Public Sub TextBox_KeyUp(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)

    If keycode = vbKeyReturn Then
        TextBox.Text = Replace(TextBox.Text, vbCrLf, "")
    End If

End Sub

Private Sub CancelButton_Click()
    
    'canceled message
    MsgBox ("Group update canceled. Previous update maintained."), , "Update Canceled"
    
    'hide form
    Me.Hide
    boolCanceled = True
    
    'call snsearch
    Call SNSearchBox
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Me.Hide
        boolCanceled = True
    End If

End Sub
