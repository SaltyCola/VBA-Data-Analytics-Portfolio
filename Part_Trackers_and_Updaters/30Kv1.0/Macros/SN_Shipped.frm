VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_Shipped 
   Caption         =   "Shipped Engine Set"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "SN_Shipped.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_Shipped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub ButtonConfirm_Click()

    Dim shpdlistInt As Double
    Dim shpdRng As Range
    Dim shpdCell As Range
    Dim colCell As Range
    Dim shpdCutCol() As Double
    
    ReDim shpdCutCol(20)
    Set shpdRng = Worksheets("NEO 5322121").Range("6:6")
    
    'iterate listbox
    For shpdlistInt = 0 To (ListBoxSN.ListCount - 1)
        'search tracker for sn
        For Each shpdCell In shpdRng
            'SN found
            If shpdCell.Value = ListBoxSN.List(shpdlistInt) Then
                
                'goto column
                Application.Goto shpdCell, Scroll:=True
                
                'change engine set number and count
                Worksheets("NEO 5322121").Cells(5, shpdCell.Column).Value = (shpdlistInt + 1)
                Worksheets("NEO 5322121").Cells(1, shpdCell.Column).Value = TextBoxEngSet.Value
                
                'iterate column to green out white cells
                For Each colCell In Worksheets("NEO 5322121").Range(Cells(7, shpdCell.Column), Cells(43, shpdCell.Column))
                    If colCell.Interior.Color = clrBlank Then
                        colCell.Value = Date
                        colCell.Interior.Color = RGB(146, 208, 80)
                    End If
                Next colCell
                
                'save column numbers
                shpdCutCol(shpdlistInt + 1) = shpdCell.Column
                
                ''move column to shipped tab
                'shpdCutCol = shpdCell.Column
                'shpdCell.EntireColumn.Cut Worksheets("Shipped").Columns(shpdFinalBlackLine + 1)
                'Worksheets("NEO 5322121").Columns(shpdCutCol).Delete
                'Application.Goto Worksheets("Shipped").Cells(6, (shpdFinalBlackLine + 1)), Scroll:=True
                'shpdFinalBlackLine = shpdFinalBlackLine + 1
                
                'last listbox entry proceedures
                If shpdlistInt = (ListBoxSN.ListCount - 1) Then
                    Worksheets("NEO 5322121").Application.Union(Columns(shpdCutCol(1)), Columns(shpdCutCol(2)), Columns(shpdCutCol(3)), Columns(shpdCutCol(4)), Columns(shpdCutCol(5)), Columns(shpdCutCol(6)), Columns(shpdCutCol(7)), Columns(shpdCutCol(8)), Columns(shpdCutCol(9)), Columns(shpdCutCol(10)), Columns(shpdCutCol(11)), Columns(shpdCutCol(12)), Columns(shpdCutCol(13)), Columns(shpdCutCol(14)), Columns(shpdCutCol(15)), Columns(shpdCutCol(16)), Columns(shpdCutCol(17)), Columns(shpdCutCol(18)), Columns(shpdCutCol(19)), Columns(shpdCutCol(20))).Copy Worksheets("Shipped").Columns(shpdFinalBlackLine + 1)
                    Worksheets("NEO 5322121").Application.Union(Columns(shpdCutCol(1)), Columns(shpdCutCol(2)), Columns(shpdCutCol(3)), Columns(shpdCutCol(4)), Columns(shpdCutCol(5)), Columns(shpdCutCol(6)), Columns(shpdCutCol(7)), Columns(shpdCutCol(8)), Columns(shpdCutCol(9)), Columns(shpdCutCol(10)), Columns(shpdCutCol(11)), Columns(shpdCutCol(12)), Columns(shpdCutCol(13)), Columns(shpdCutCol(14)), Columns(shpdCutCol(15)), Columns(shpdCutCol(16)), Columns(shpdCutCol(17)), Columns(shpdCutCol(18)), Columns(shpdCutCol(19)), Columns(shpdCutCol(20))).Delete
                    shpdFinalBlackLine = shpdFinalBlackLine + 21
                    Worksheets("Shipped").Columns(shpdFinalBlackLine).EntireColumn.Interior.Color = RGB(0, 0, 0)
                    ReDim shpdCutCol(20)
                End If
                
            End If
        Next shpdCell
    Next shpdlistInt

    'make first tracker engine set increase
    Worksheets("NEO 5322121").Cells(1, 3).Value = (TextBoxEngSet.Value + 1)

    'redo engine sets and go back to main menu
    Me.Hide
    'delete black columns first
    Call Mod_WIPUpdater.DeleteBlackColumns
    'finalize engine sets
    Call Mod_WIPUpdater.EngineSetHandler
    'go to main menu
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub ButtonRemoveEntry_Click()

    'remove selected item
    If (ListBoxSN.ListCount > 0) Then
        ListBoxSN.RemoveItem (ListBoxSN.ListIndex)
        'decrease engine set count
        TextBoxCount.Value = (TextBoxCount.Value - 1)
        'scroll to and select last item in list
        If (ListBoxSN.ListCount > 0) Then
            ListBoxSN.Selected(frmSNShipped.ListBoxSN.ListCount - 1) = True
        End If
        'disable confirm button and enable SN text box
        If (frmSNShipped.TextBoxCount.Value < 20) Then
            frmSNShipped.ButtonConfirm.Enabled = False
        End If
    End If
    
    'set focus back to text box
    Me.Hide
    TextBoxSN.Enabled = True
    TextBoxSN.SetFocus
    Me.Show

End Sub

Private Sub TextBoxSN_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)

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
        Set snRng = Worksheets("NEO 5322121").Range("6:6")
    
        'Define variables
        trgtStr = TextBoxSN.Value
        
        'initialize snml type
        intSNMLType = 4
        
        'Length / Format Error Msgs
        If (Len(trgtStr) < 4) Or (Len(trgtStr) > 5) Then
            MsgBox ("Please enter the correct serial number format. (i.e. J0101 or 0101)"), , "Length Error"
            'clear out textbox
            Me.Hide
            frmSNShipped.TextBoxSN.Value = ""
            frmSNShipped.TextBoxSN.SetFocus
            Me.Show
            Exit Sub
        ElseIf (Len(trgtStr) = 5) And (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid(trgtStr, 1, 1))) = 0) Then
            MsgBox ("The first character of a five character entry must be a letter. (i.e. J0101)"), , "Format Error"
            'clear out textbox
            Me.Hide
            frmSNShipped.TextBoxSN.Value = ""
            frmSNShipped.TextBoxSN.SetFocus
            Me.Show
            Exit Sub
        ElseIf (InStr("0123456789", Mid(Right(trgtStr, 4), 1, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 2, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 3, 1)) = 0) Or (InStr("0123456789", Mid(Right(trgtStr, 4), 4, 1)) = 0) Then
            MsgBox ("The final 4 characters must be numbers. (i.e. J0101 or 0101)"), , "Format Error"
            'clear out textbox
            Me.Hide
            frmSNShipped.TextBoxSN.Value = ""
            frmSNShipped.TextBoxSN.SetFocus
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
            Call SNMatchList
            
        End If
        '==========================================================================='
        
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        boolCanceled = True
    End If

End Sub
