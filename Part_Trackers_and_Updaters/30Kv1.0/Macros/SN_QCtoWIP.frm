VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_QCtoWIP 
   Caption         =   "QC to WIP"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "SN_QCtoWIP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_QCtoWIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub SN_QCtoWIP_TextBox_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)
    
    If keycode = vbKeyReturn Then
        
        '==============================Search Tracker==============================='
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
        Set snRng = Worksheets("Quality Clinic").Range("6:6")
    
        'Define variables
        trgtStr = SN_QCtoWIP_TextBox.Value
        
        'initialize snml type
        intSNMLType = 5
        
        'Length / Format Error Msgs
        If (Len(trgtStr) < 4) Or (Len(trgtStr) > 5) Then
            Me.Hide
            MsgBox ("Please enter the correct serial number format. (i.e. J0101 or 0101)"), , "Length Error"
            intError = True
            Exit Sub
        ElseIf (Len(trgtStr) = 5) And (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(trgtStr, 1, 1))) = 0) Then
            Me.Hide
            MsgBox ("The first character of a five character entry must be a letter. (i.e. J0101)"), , "Format Error"
            intError = True
            Exit Sub
        ElseIf (InStr("0123456789", Mid(Right(trgtStr, 4), 1, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 2, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 3, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 4, 1)) = 0) Then
            Me.Hide
            MsgBox ("The final 4 characters must be numbers. (i.e. J0101 or 0101)"), , "Format Error"
            intError = True
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
            
            'Call Match list
            Me.Hide
            Call SNMatchList
            
        End If
        '==========================================================================='
        
    End If

End Sub

Private Sub ButtonMovetoWIP_Click()

    Dim colRedLineQCtoWIP As Double
    Dim wipcll As Range
    Dim colQC As Double
    
    'grab active SN's Column #
    colQC = ActiveCell.Column
    
    'find red column
    For Each wipcll In Worksheets("NEO 5322121").Range("6:6")
        If wipcll.Interior.Color = RGB(255, 0, 0) Then
            colRedLineQCtoWIP = wipcll.Column
            Exit For
        End If
    Next wipcll
    
    'cut and paste qc SN
    Worksheets("NEO 5322121").Columns(colRedLineQCtoWIP).Insert
    Worksheets("Quality Clinic").Columns(colQC).Cut Worksheets("NEO 5322121").Columns(colRedLineQCtoWIP)
    Worksheets("Quality Clinic").Columns(colQC).Delete
    Application.Goto Worksheets("NEO 5322121").Cells(6, colRedLineQCtoWIP), Scroll:=True
    
    'hide and call sninfopage
    Me.Hide
    Call SNInfoPage

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        boolCanceled = True
    End If

End Sub
