VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_AsBuilt 
   Caption         =   "As Built Updater"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "SN_AsBuilt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_AsBuilt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub TextBoxSN_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)

    If keycode = vbKeyReturn Then
        
        '==============================Search Tracker==============================='
        Dim trgtStrAB As String
        Dim snRngAB As Range
        Dim snAB As Range
        Dim snColAB As Long
        Dim snRowAB As Long
        Dim snTxtRawAB As String
        Dim snTxtAB As String
        Dim snTxtPrefixAB
        Dim snTxtUAB As String
        Dim snTxtLAB As String
    
        'Define snRngAB
        Set snRngAB = Worksheets("NEO 5322121").Range("6:6")
    
        'Define variables
        trgtStrAB = frmSNAsBuilt.TextBoxSN.Value
        
        'initialize snml type
        intSNMLType = 3
        
        'Length / Format Error Msgs
        If (Len(trgtStrAB) < 4) Or (Len(trgtStrAB) > 5) Then
            Me.Hide
            MsgBox ("Please enter the correct serial number format. (i.e. J0101 or 0101)"), , "Length Error"
            intError = True
            boolMaintoAB = True
            'call snAsBuilt
            Call SNAsBuilt
            Exit Sub
        ElseIf (Len(trgtStrAB) = 5) And (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(trgtStrAB, 1, 1))) = 0) Then
            Me.Hide
            MsgBox ("The first character of a five character entry must be a letter. (i.e. J0101)"), , "Format Error"
            intError = True
            boolMaintoAB = True
            'call snAsBuilt
            Call SNAsBuilt
            Exit Sub
        ElseIf (InStr("0123456789", Mid(Right(trgtStrAB, 4), 1, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStrAB, 4), 2, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStrAB, 4), 3, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStrAB, 4), 4, 1)) = 0) Then
            Me.Hide
            MsgBox ("The final 4 characters must be numbers. (i.e. J0101 or 0101)"), , "Format Error"
            intError = True
            boolMaintoAB = True
            'call snAsBuilt
            Call SNAsBuilt
            Exit Sub
        
        'serial number correct format
        Else
            'search sn row in tracker
            For Each snAB In snRngAB
                snTxtRawAB = snAB.Value
                snRowAB = snAB.Row
                snColAB = snAB.Column
                'only look at SN's
                If Len(snTxtRawAB) > 5 Then
                    snTxtAB = Right(snTxtRawAB, (Len(snTxtRawAB) - 5))
                    snTxtPrefixAB = Mid(snTxtAB, 1, 1)
                    snTxtUAB = UCase(snTxtPrefixAB) & Mid(snTxtAB, 2)
                    snTxtLAB = LCase(snTxtPrefixAB) & Mid(snTxtAB, 2)
                    
                    'On serial number found:
                    If ((trgtStrAB = snTxtUAB) Or (trgtStrAB = Right(snTxtUAB, 4))) Or ((trgtStrAB = snTxtLAB) Or (trgtStrAB = Right(snTxtLAB, 4))) Then
                        'first sn match
                        If SNMLArrayCnt = 0 Then
                            'populate matchlist array
                            SNMLArray(UBound(SNMLArray)) = snAB.Value
                            SNMLArrayCnt = SNMLArrayCnt + 1
                        'every match after the first
                        ElseIf SNMLArrayCnt > 0 Then
                            'redimension array
                            ReDim Preserve SNMLArray(SNMLArrayCnt)
                            'populate matchlist array
                            SNMLArray(UBound(SNMLArray)) = snAB.Value
                            SNMLArrayCnt = SNMLArrayCnt + 1
                        End If
                    
                    End If
                End If
            Next snAB
            
            'from As Built boolean
            boolAsBuilt = True
            boolMaintoAB = False
            
            'Call Match list
            Call SNMatchList
            
        End If
        '==========================================================================='
        
    End If

End Sub

Private Sub ToggleAccepted_Click()

    'if button toggled
    If ToggleAccepted.Value = True Then
        ActiveCell.Interior.Color = RGB(146, 208, 80)
        ToggleRejected.Enabled = False
    'if button untoggled
    ElseIf ToggleAccepted.Value = False Then
        ActiveCell.Interior.Color = clrBlank
        ToggleRejected.Enabled = True
    End If

End Sub

Private Sub ToggleRejected_Click()

    'if button toggled
    If ToggleRejected.Value = True Then
        ActiveCell.Interior.Color = RGB(255, 0, 0)
        ToggleAccepted.Enabled = False
    'if button untoggled
    ElseIf ToggleRejected.Value = False Then
        ActiveCell.Interior.Color = clrBlank
        ToggleAccepted.Enabled = True
    End If

End Sub

Private Sub TextBoxDate_Change()

    ActiveCell.Value = TextBoxDate.Value

End Sub

Private Sub SN_AsBuilt_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        intError = False
    End If

End Sub
