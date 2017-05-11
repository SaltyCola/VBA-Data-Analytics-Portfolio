VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_NewSN 
   Caption         =   "New Serial Numbers"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2880
   OleObjectBlob   =   "SN_NewSN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_NewSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub Userform_Activate()

    'initialize SN Template Arrays
    ReDim arraySNTemplateVal(56)
    ReDim arraySNTemplateClr(56)
    Dim i As Double
    For i = 1 To 56
        If i >= 7 And i <= 42 Then: arraySNTemplateVal(i) = "=WORKDAY($%&!*!&%$" & (i + 1) & ",'NEO 5322121'!$A" & i & ",'Holiday Schedule'!$B$2:$B$30)"
        If i = 43 Then: arraySNTemplateVal(i) = Str(Date)
        If i = 4 Then: arraySNTemplateClr(i) = RGB(197, 217, 241)
        If i <> 4 Then: arraySNTemplateClr(i) = clrBlank
    Next i

End Sub

Private Sub TextBoxSNEntry_Change()

    'if first character is a letter and last 4 are numbers and only 5 characters, enable create SN button
    If (Len(TextBoxSNEntry.Value) = 5) Then
        If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(TextBoxSNEntry.Value, 1))) <> 0) And (InStr("0123456789", Mid(Right(TextBoxSNEntry.Value, 4), 1, 1)) <> 0) And (InStr("0123456789", Mid(Right(TextBoxSNEntry.Value, 4), 2, 1)) <> 0) And (InStr("0123456789", Mid(Right(TextBoxSNEntry.Value, 4), 3, 1)) <> 0) And (InStr("0123456789", Mid(Right(TextBoxSNEntry.Value, 4), 4, 1)) <> 0) Then
            ButtonCreate.Enabled = True
        '5 chars but not correct types
        Else
            ButtonCreate.Enabled = False
        End If
    'any length besides 5 chars
    Else
        ButtonCreate.Enabled = False
    End If

End Sub

Private Sub TextBoxSNEntry_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal shift As Integer)

    If keycode = vbKeyReturn Then
        'if create button enabled
        If ButtonCreate.Enabled = True Then
            ButtonCreate.SetFocus
        'if create button disabled
        ElseIf ButtonCreate.Enabled = False Then
            MsgBox ("Please enter the correct serial number format. (i.e. J0101, K1010, etc.)"), , "Format Error"
            Me.Hide
            TextBoxSNEntry.SetFocus
            Me.Show
        End If
    End If

End Sub

Private Sub ToggleUnlockPrefix_Click()

    'enable or disable prefix textbox
    If ToggleUnlockPrefix.Value = False Then
        TextBoxSNPrefix.Enabled = True
    ElseIf ToggleUnlockPrefix.Value = True Then
        TextBoxSNPrefix.Enabled = False
    End If

End Sub

Private Sub ButtonCreate_Click()

    'make first char UCase
    TextBoxSNEntry.Value = UCase(TextBoxSNEntry.Value)
    
    'grab SN Text
    strNewSN = TextBoxSNPrefix.Value & TextBoxSNEntry.Value
    
    'apply to index 6 of new sn value array
    arraySNTemplateVal(6) = strNewSN
    
    'find red line column
    Set redlineRng = Worksheets("NEO 5322121").Range("6:6")
    For Each redlineCell In redlineRng
        If redlineCell.Interior.Color = RGB(255, 0, 0) Then
            redlineInt = redlineCell.Column
            replaceInt = InStr(redlineCell.EntireColumn.Address, ":")
            redlineColAddress = Replace(Left(redlineCell.EntireColumn.Address, (replaceInt - 1)), "$", "")
            Exit For
        End If
    Next redlineCell
    
    'fix address in arrays
    Dim i As Double
    For i = 7 To 42 'ignore last row (date)
        arraySNTemplateVal(i) = Replace(arraySNTemplateVal(i), "$%&!*!&%$", redlineColAddress)
    Next i
    
    'insert new column before red line
    Worksheets("NEO 5322121").Columns(redlineInt).Insert
    
    'apply all new sn template values and colors
    Dim j As Double
    For j = 1 To 56
        Worksheets("NEO 5322121").Cells(j, redlineInt).Value = arraySNTemplateVal(j)
        Worksheets("NEO 5322121").Cells(j, redlineInt).Interior.Color = arraySNTemplateClr(j)
    Next j
    
    'fix white cell dates and colors
    Dim ff As Double
    Dim f As Range
    Dim ldTime As Double
    For ff = 43 To 7 Step -1
        Set f = Worksheets("NEO 5322121").Cells(ff, redlineInt)
        If Not (f.Interior.Color = RGB(146, 208, 80)) And Not (f.Interior.Color = RGB(79, 98, 40)) And Not (f.Interior.Color = RGB(196, 215, 155)) And Not (f.Interior.Color = RGB(0, 176, 80)) And Not (f.Interior.Color = RGB(255, 192, 0)) And Not (f.Interior.Color = RGB(146, 205, 220)) And Not (f.Interior.Color = RGB(255, 0, 0)) And Not (f.Interior.Color = RGB(0, 0, 0)) Then
            'set lead time variable
            If Worksheets("NEO 5322121").Cells(f.Row, 1).Value = 0.5 Then
                ldTime = 0
            Else
                ldTime = Worksheets("NEO 5322121").Cells(f.Row, 1).Value
            End If
            'todays date for bottom row
            If ff = 43 Then
                f.Value = Date
            Else
                f.Value = f.Offset(1, 0).Value + ldTime
            End If
            f.Interior.Color = RGB(255, 255, 255)
        End If
    Next ff
    
    'scroll
    Application.Goto Worksheets("NEO 5322121").Columns(redlineInt), Scroll:=True
    
    'send focus to textbox
    TextBoxSNEntry.SetFocus
    
    'reset variables
    redlineInt = 0
    replaceInt = 0
    redlineColAddress = 0
    Call Userform_Activate

End Sub

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub
