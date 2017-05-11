Attribute VB_Name = "Mod_WIPUpdater"
Option Explicit

Public Sub SNNewSN()
    
    'load SN_NewSN
    Set frmSNNewSN = New SN_NewSN
    frmSNNewSN.Show

End Sub

Public Sub SNSearchBox()
'Serial Number Search Box

'reference line for repeates after error
lineSearchBox:
    
    'Initialize variables
    boolCanceled = False
    intError = False
    boolTog = False
    boolTogsAllowed = False
    boolRTOTog = False
    boolSNDTGsnNotFound = False
    txtBoxTxt = ""
    txtBoxClr = 0
    togClr = 0
    TodaysDate = Str(Date)
    SNIPRow = 0
    SNIPTog = "0"
    SNIPSearchCurrentRow = 44
    ReDim SNMLArray(0)
    SNMLArrayCnt = 0
    intSNMLType = 0
    clrBlank = RGB(255, 255, 255)
    
    'initialize arrays
    ReDim arraySNBackupVal(56)
    ReDim arraySNBackupClr(56)
    ReDim arraySNUpdatedVal(56)
    ReDim arraySNUpdatedClr(56)
    ReDim arrayUpdatesOnlyVal(56)
    ReDim arrayUpdatesOnlyClr(56)
        '56th entry is SN column before waterfall!!!!
    ReDim arrayWaterfallVal(57)
        '============================================
    ReDim arrayWaterfallClr(56)
    
    '=================Updates Only Arrays Initializer================='
    Dim arr As Double
    For arr = 1 To 56
        arrayUpdatesOnlyVal(arr) = "!===N/A===!"
        arrayUpdatesOnlyClr(arr) = 102030405
    Next arr
    '=================Updates Only Arrays Initializer================='
    
    'reset intError and boolCanceled
    intError = False
    boolCanceled = False

    'Load SN_Search form
    Set frmSNSearch = New SN_Search
    frmSNSearch.Show
    
    'On Error rerun form
    If intError Then
        GoTo lineSearchBox
    End If

End Sub

Public Sub SNQCtoWIP()
'Serial Number Search Box

'reference line for repeates after error
lineSearchBoxQC:
    
    'Initialize variables
    boolCanceled = False
    intError = False
    boolTog = False
    boolTogsAllowed = False
    boolRTOTog = False
    boolSNDTGsnNotFound = False
    txtBoxTxt = ""
    txtBoxClr = 0
    togClr = 0
    TodaysDate = Str(Date)
    SNIPRow = 0
    SNIPTog = "0"
    SNIPSearchCurrentRow = 44
    ReDim SNMLArray(0)
    SNMLArrayCnt = 0
    intSNMLType = 0
    clrBlank = RGB(255, 255, 255)
    
    'initialize arrays
    ReDim arraySNBackupVal(56)
    ReDim arraySNBackupClr(56)
    ReDim arraySNUpdatedVal(56)
    ReDim arraySNUpdatedClr(56)
    ReDim arrayUpdatesOnlyVal(56)
    ReDim arrayUpdatesOnlyClr(56)
        '56th entry is SN column before waterfall!!!!
    ReDim arrayWaterfallVal(57)
        '============================================
    ReDim arrayWaterfallClr(56)
    
    '=================Updates Only Arrays Initializer================='
    Dim arr As Double
    For arr = 1 To 56
        arrayUpdatesOnlyVal(arr) = "!===N/A===!"
        arrayUpdatesOnlyClr(arr) = 102030405
    Next arr
    '=================Updates Only Arrays Initializer================='
    
    'reset intError and boolCanceled
    intError = False
    boolCanceled = False

    'Load SN_Search form
    Set frmQCtoWIP = New SN_QCtoWIP
    frmQCtoWIP.SN_QCtoWIP_TextBox.SetFocus
    frmQCtoWIP.Show
    
    'On Error rerun form
    If intError Then
        GoTo lineSearchBoxQC
    End If

End Sub

Public Sub SNAsBuilt()

lineAsBuilt:
    
    'Initialize variables
    boolCanceled = False
    intError = False
    boolTog = False
    boolTogsAllowed = False
    boolRTOTog = False
    boolSNDTGsnNotFound = False
    txtBoxTxt = ""
    txtBoxClr = 0
    togClr = 0
    TodaysDate = Str(Date)
    SNIPRow = 0
    SNIPTog = "0"
    SNIPSearchCurrentRow = 44
    ReDim SNMLArray(0)
    SNMLArrayCnt = 0
    intSNMLType = 0
    clrBlank = RGB(255, 255, 255)
    boolAsBuilt = False
    
    'initialize arrays
    ReDim arraySNBackupVal(56)
    ReDim arraySNBackupClr(56)
    ReDim arraySNUpdatedVal(56)
    ReDim arraySNUpdatedClr(56)
    ReDim arrayUpdatesOnlyVal(56)
    ReDim arrayUpdatesOnlyClr(56)
        '56th entry is SN column before waterfall!!!!
    ReDim arrayWaterfallVal(57)
        '============================================
    ReDim arrayWaterfallClr(56)
    
    '=================Updates Only Arrays Initializer================='
    Dim arr As Double
    For arr = 1 To 56
        arrayUpdatesOnlyVal(arr) = "!===N/A===!"
        arrayUpdatesOnlyClr(arr) = 102030405
    Next arr
    '=================Updates Only Arrays Initializer================='
    
    'reset intError and boolCanceled
    intError = False
    boolCanceled = False

    'Load SN_AsBuilt form if not focus is not coming from Match List
    If boolMaintoAB Then
        Set frmSNAsBuilt = New SN_AsBuilt
        frmSNAsBuilt.ToggleAccepted.BackColor = RGB(146, 208, 80)
        frmSNAsBuilt.ToggleRejected.BackColor = RGB(255, 0, 0)
        frmSNAsBuilt.Show
        boolMaintoAB = False
    End If
    
    'fill in textboxes
    If boolSNMLtoAB Then
        frmSNAsBuilt.Caption = snTitleCell.Value
        frmSNAsBuilt.TextBoxSN.Value = ""
        frmSNAsBuilt.TextBoxDate.Enabled = True
        frmSNAsBuilt.TextBoxDate.Value = Worksheets("NEO 5322121").Cells(55, abCol).Value
        'initialize toggle buttons
        If Worksheets("NEO 5322121").Cells(55, abCol).Interior.Color = RGB(146, 208, 80) Then
            frmSNAsBuilt.ToggleAccepted.Enabled = True
            frmSNAsBuilt.ToggleAccepted.Value = True
            frmSNAsBuilt.ToggleRejected.Value = False
        ElseIf Worksheets("NEO 5322121").Cells(55, abCol).Interior.Color = RGB(255, 0, 0) Then
            frmSNAsBuilt.ToggleRejected.Enabled = True
            frmSNAsBuilt.ToggleRejected.Value = True
            frmSNAsBuilt.ToggleAccepted.Value = False
        Else
            frmSNAsBuilt.ToggleAccepted.Enabled = True
            frmSNAsBuilt.ToggleAccepted.Value = False
            frmSNAsBuilt.ToggleRejected.Enabled = True
            frmSNAsBuilt.ToggleRejected.Value = False
        End If
        boolSNMLtoAB = False
    End If
    
    'On Error rerun form
    If intError Then
        GoTo lineAsBuilt
    End If

End Sub

Public Sub SNSlowPartsAnalyzer()

    Dim snSlowRng As Range
    Dim snSlowCell As Range
    Dim snSlowCount As Double
    
    'initialize variables
    boolCanceled = False
    snSlowCount = 0
    
    'define range
    Set snSlowRng = Worksheets("NEO 5322121").Range("6:6")
    
    'load form
    Set frmSNSlowParts = New SN_SlowParts
    frmSNSlowParts.ToggleSlowPart.BackColor = RGB(244, 158, 228)
    frmSNSlowParts.TextBoxToday.Value = Date
    
    'search tracker to find possible slow movers
    For Each snSlowCell In snSlowRng
        
        'show progress
        Application.Goto snSlowCell, Scroll:=True
    
        'exit for loop on red line found
        If snSlowCell.Interior.Color = RGB(255, 0, 0) Then
            'go to beginning of tracker
            Application.Goto Worksheets("NEO 5322121").Cells(6, 3), Scroll:=True
            Exit For
        End If
        
        'if last date seen is >= 7 days ago, add to listbox (ignore already slow parts)
        If (Worksheets("NEO 5322121").Cells(52, snSlowCell.Column).Value <= (Date - 7)) And Not (Worksheets("NEO 5322121").Cells(52, snSlowCell.Column).Value = "") And Not (snSlowCell.Interior.Color = RGB(244, 158, 228)) And Not (IsEmpty(snSlowCell)) And (snSlowCell.Column > 2) Then
            frmSNSlowParts.ListBox1.AddItem (snSlowCell.Value)
            snSlowCount = snSlowCount + 1
        End If
    
    Next snSlowCell
    
    'give total number found
    frmSNSlowParts.Caption = frmSNSlowParts.Caption & ": " & snSlowCount & " Found."
    
    'show form
    frmSNSlowParts.Show

End Sub

Public Sub SNShipped()

    'reference line for repeates after error
lineShipped:
    
    'Initialize variables
    boolCanceled = False
    intError = False
    boolTog = False
    boolTogsAllowed = False
    boolRTOTog = False
    boolSNDTGsnNotFound = False
    txtBoxTxt = ""
    txtBoxClr = 0
    togClr = 0
    TodaysDate = Str(Date)
    SNIPRow = 0
    SNIPTog = "0"
    SNIPSearchCurrentRow = 44
    ReDim SNMLArray(0)
    SNMLArrayCnt = 0
    intSNMLType = 0
    clrBlank = RGB(255, 255, 255)
    
    'initialize arrays
    ReDim arraySNBackupVal(56)
    ReDim arraySNBackupClr(56)
    ReDim arraySNUpdatedVal(56)
    ReDim arraySNUpdatedClr(56)
    ReDim arrayUpdatesOnlyVal(56)
    ReDim arrayUpdatesOnlyClr(56)
        '56th entry is SN column before waterfall!!!!
    ReDim arrayWaterfallVal(57)
        '============================================
    ReDim arrayWaterfallClr(56)
    
    '=================Updates Only Arrays Initializer================='
    Dim arr As Double
    For arr = 1 To 56
        arrayUpdatesOnlyVal(arr) = "!===N/A===!"
        arrayUpdatesOnlyClr(arr) = 102030405
    Next arr
    '=================Updates Only Arrays Initializer================='

    'Load SN_Shipped form
    'if first time run
    If Not boolSNMLtoShipped Then
        Set frmSNShipped = New SN_Shipped
        'look up next engine set
        Dim curengsetRng As Range
        Dim curengsetCell As Range
        Set curengsetRng = Worksheets("Shipped").Range("1:1")
        For Each curengsetCell In curengsetRng
            'if more current engine set found
            If IsNumeric(curengsetCell.Value) And (curengsetCell.Value > 0) Then
                CurrentEngineSet = curengsetCell.Value + 1
            End If
            'black line found
            If curengsetCell.Interior.Color = RGB(0, 0, 0) Then
                shpdFinalBlackLine = curengsetCell.Column
            End If
        Next curengsetCell
        frmSNShipped.TextBoxEngSet.Value = CurrentEngineSet
        frmSNShipped.ButtonConfirm.Enabled = False
        frmSNShipped.Show
    End If
    
    'if coming from snml
    If boolSNMLtoShipped Then
        'check for existence in listbox
        Dim chkInt As Double
        For chkInt = 0 To (frmSNShipped.ListBoxSN.ListCount - 1)
            If (frmSNShipped.ListBoxSN.List(chkInt) = snTitleCell.Value) Then
                MsgBox "That Serial Number is already in the list!"
                GoTo lineNotAdded
            ElseIf (boolShippedDoNotAdd) Then
                GoTo lineNotAdded
            End If
        Next chkInt
        'add SN to list
        frmSNShipped.ListBoxSN.AddItem (snTitleCell.Value)
        'increase Engine Set Count
        frmSNShipped.TextBoxCount.Value = (frmSNShipped.TextBoxCount.Value + 1)
lineNotAdded:
        'scroll to and select newest item
        frmSNShipped.ListBoxSN.Selected(frmSNShipped.ListBoxSN.ListCount - 1) = True
        'clear out textbox
        frmSNShipped.Hide
        frmSNShipped.TextBoxSN.Value = ""
        frmSNShipped.TextBoxSN.SetFocus
        'enable confirm button
        If (Int(frmSNShipped.TextBoxCount.Value) >= 20) Then
            frmSNShipped.TextBoxSN.Enabled = False
            frmSNShipped.ButtonConfirm.Enabled = True
        End If
        'unhide form
        frmSNShipped.Show
        
        'reset booleans
        boolSNMLtoShipped = False
        boolShippedDoNotAdd = False
    End If
    
    'if boolCanceled = true exit sub
    If boolCanceled Then
        Exit Sub
    End If
    
    'On Error rerun form
    'If intError Then
    '    GoTo lineShipped
    'End If

End Sub

Public Sub SNMatchList()

    'Move on if forms not closed
    If Not boolCanceled Then
    
        'Load SN_MatchList
        Set frmSNMatchList = New SN_MatchList
        
        'populate combobox with list array and display first entry
        frmSNMatchList.ComboBox1.List = SNMLArray
        frmSNMatchList.ComboBox1.ListIndex = 0
        
        'check if combobox is empty
        If (frmSNMatchList.ComboBox1.ListCount = 1) And (frmSNMatchList.ComboBox1.Value = "") Then
            
            'Search Error Msg for WIP
            boolSNDTGsnNotFound = True
            If intSNMLType <> 5 Then
                MsgBox ("Serial number not found in Active WIP..."), , "Search Error"
            ElseIf intSNMLType = 5 Then
                MsgBox ("Serial number not found in Quality Clinic..."), , "Search Error"
            End If
            
            'coming from sn search
            If intSNMLType = 1 Then
                'call snsearch
                Call SNSearchBox
            'coming from sn duplicate to group
            ElseIf intSNMLType = 2 Then
                'clear out sn duplicate textbox
                frmSNDuplicate.TextBox.Value = ""
                'call snduplicatetogroup
                Call SNDuplicateToGroup
            'coming from As Built
            ElseIf intSNMLType = 3 Then
                'call snAsBuilt
                Call SNAsBuilt
            'coming from SN Shipped
            ElseIf intSNMLType = 4 Then
                'send focus to textbox
                frmSNShipped.Hide
                frmSNShipped.TextBoxSN.SetFocus
                frmSNShipped.Show
            ElseIf intSNMLType = 5 Then
                'call SNQCtoWIP
                Call SNQCtoWIP
            End If
        
        
        'if combobox has only one match
        ElseIf (frmSNMatchList.ComboBox1.ListCount = 1) And Not (frmSNMatchList.ComboBox1.Value = "") Then
            
            Dim snRng As Range
            Dim sn As Range
            Dim snCol As Long
            Dim snRow As Long
            
            'Define snRng
            If intSNMLType <> 5 Then
                Set snRng = Worksheets("NEO 5322121").Range("6:6")
            ElseIf intSNMLType = 5 Then
                Set snRng = Worksheets("Quality Clinic").Range("6:6")
            End If
            
            'search sn row in tracker
            For Each sn In snRng
                snRow = sn.Row
                snCol = sn.Column
    
                'If match
                If sn.Value = frmSNMatchList.ComboBox1.Value Then
            
                    'go to found sn
                    If intSNMLType <> 5 Then
                        Application.Goto Worksheets("NEO 5322121").Cells(snRow, snCol), Scroll:=True
                    ElseIf intSNMLType = 5 Then
                        Application.Goto Worksheets("Quality Clinic").Cells(snRow, snCol), Scroll:=True
                    End If
                    
                    'Define current SN cell
                    Set snTitleCell = ActiveCell
                    
                    'if from as built, send to as built row
                    If boolAsBuilt Then
                        Application.Goto Worksheets("NEO 5322121").Cells(55, snCol), Scroll:=True
                    End If
                    
                    'clear combobox
                    frmSNMatchList.ComboBox1.Clear
                    
                    'if coming from snsearch, send to sn info page
                    If intSNMLType = 1 Then
                        Call SNInfoPage
                        Exit Sub
                    'if coming from snduplicate, send back to snduplicate
                    ElseIf intSNMLType = 2 Then
                        'set first run to false
                        boolSNDTGFirstRun = False
                        'clear out sn duplicate textbox
                        frmSNDuplicate.TextBox.Value = ""
                        'call snduplicatetogroup
                        Call SNDuplicateToGroup
                        Exit Sub
                    'if coming from as built, send back to as built
                    ElseIf intSNMLType = 3 Then
                        'set boolean and abCol for SNML to AB
                        boolSNMLtoAB = True
                        abCol = snCol
                        'call snAsBuilt
                        Call SNAsBuilt
                        Exit Sub
                    'if coming from snshipped, send back to snshipped
                    ElseIf intSNMLType = 4 Then
                        'set boolean
                        boolSNMLtoShipped = True
                        'call snShipped
                        Call SNShipped
                        Exit Sub
                    'if coming from snQCtoWIP, send focus to Move to WIP button
                    ElseIf intSNMLType = 5 Then
                        frmQCtoWIP.ButtonMovetoWIP.Enabled = True
                        frmQCtoWIP.ButtonMovetoWIP.SetFocus
                        frmQCtoWIP.Show
                        Exit Sub
                    End If
                    
                End If
            Next sn
        
        
        'if combobox has more than one match
        ElseIf (frmSNMatchList.ComboBox1.ListCount > 1) Then
            
            'show form
            frmSNMatchList.Show
        
        End If
    End If

End Sub

'Serial Number Info Page
Public Sub SNInfoPage()
    
    Dim r As Range

    'Move on if forms not closed
    If Not boolCanceled Then
        
        'Load SN_InfoPage
        Set frmSNInfoPage = New SN_InfoPage
        
        'if from QCtoWIP disable group update button
        If intSNMLType = 5 Then
            frmSNInfoPage.UpdateGroupButton.Enabled = False
            'get correct SN (WORKAROUND)
            For Each r In Worksheets("NEO 5322121").Range("6:6")
                If r.Offset(0, 1).Interior.Color = RGB(255, 0, 0) Then
                    Set snTitleCell = r
                    frmSNInfoPage.Caption = r.Value
                    Exit For
                End If
            Next r
        End If
        
        'InfoPage initializers
        Call SNIPCurrentInfo
        Call SNIPColorOrg
        
        'show form
        frmSNInfoPage.Show
        
    End If
    
End Sub

'Duplicate SN Update to Group of SN's
Public Sub SNDuplicateToGroup()

'Move on if forms not closed
    If Not boolCanceled Then
    
        'Initialize variables
        ReDim SNMLArray(0)
        SNMLArrayCnt = 0
        intSNMLType = 0
    
        'Load SN_DuplicateToGroup form if not already open
        If boolSNDTGFirstRun Then
            Set frmSNDuplicate = New SN_DuplicateToGroup
            frmSNDuplicate.Show
        
        'add sn to listbox if already open
        ElseIf boolSNDTGFirstRun = False Then
            
            'listbox empty
            If frmSNDuplicate.ListBox1.ListCount = 0 Then
                frmSNDuplicate.Hide
                frmSNDuplicate.ListBox1.AddItem (snTitleCell.Value)
                'scroll to newest item
                frmSNDuplicate.ListBox1.Selected(frmSNDuplicate.ListBox1.ListCount - 1) = True
                frmSNDuplicate.ListBox1.Selected(frmSNDuplicate.ListBox1.ListCount - 1) = False
                    'enable confirm button
                frmSNDuplicate.ConfirmButton.Enabled = True
                frmSNDuplicate.TextBox.Value = ""
                frmSNDuplicate.Show
            
            'listbox not empty
            ElseIf frmSNDuplicate.ListBox1.ListCount > 0 Then
                'initialize variables
                Dim lstInt As Double
                boolSNDTGAdd = True
                'iterate listbox entries
                For lstInt = 0 To (frmSNDuplicate.ListBox1.ListCount - 1)
                    'if sn isn't in listbox already
                    If Not frmSNDuplicate.ListBox1.List(lstInt) = snTitleCell.Value Then
                        boolSNDTGAdd = True
                    'if sn is already in listbox
                    ElseIf frmSNDuplicate.ListBox1.List(lstInt) = snTitleCell.Value Then
                        boolSNDTGAdd = False
                        Exit For
                    End If
                Next lstInt
                
                'decide to add SN or not
                If boolSNDTGAdd Then
                    frmSNDuplicate.Hide
                    frmSNDuplicate.ListBox1.AddItem (snTitleCell.Value)
                    'scroll to newest item
                    frmSNDuplicate.ListBox1.Selected(frmSNDuplicate.ListBox1.ListCount - 1) = True
                    frmSNDuplicate.ListBox1.Selected(frmSNDuplicate.ListBox1.ListCount - 1) = False
                    'clear textbox
                    frmSNDuplicate.TextBox.Value = ""
                    frmSNDuplicate.Show
                ElseIf Not (boolSNDTGAdd) And Not (boolSNDTGsnNotFound) Then
                    frmSNDuplicate.Hide
                    MsgBox ("Serial number already in list!"), , "SN Entry Error"
                    frmSNDuplicate.TextBox.Value = ""
                    frmSNDuplicate.Show
                'SN not found in SNMatchList
                ElseIf boolSNDTGsnNotFound Then
                    frmSNDuplicate.Hide
                    boolSNDTGsnNotFound = False
                    frmSNDuplicate.TextBox.Value = ""
                    frmSNDuplicate.Show
                End If
                
            End If
        End If
        
        'clear out textbox
        frmSNDuplicate.TextBox.Value = ""
    
    End If

End Sub

Sub SNIPColorOrg()

        'InfoPage Outsourced Colors
        frmSNInfoPage.Label33.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.Label38.BackColor = RGB(197, 217, 241)
        frmSNInfoPage.Label42.BackColor = RGB(197, 217, 241)
        frmSNInfoPage.Label26.BackColor = RGB(197, 217, 241)
        frmSNInfoPage.Label44.BackColor = RGB(197, 217, 241)
        frmSNInfoPage.Label49.BackColor = RGB(197, 217, 241)
        'InfoPage Update Button Colors
        frmSNInfoPage.SearchNewButton.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.UpdateGroupButton.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.CancelButton.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.SlowButton.BackColor = RGB(244, 158, 228)
        frmSNInfoPage.ESTButton.BackColor = RGB(255, 255, 0)
        'InfoPage Toggle Button Colors
            'Green (dark, light, and bright intermixed)
        frmSNInfoPage.ToggleButton1R1.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R2.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR3.BackColor = RGB(196, 215, 155)
        frmSNInfoPage.ToggleButton1R3.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R4.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R5.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R6.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R7.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R8.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R9.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R10.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR11.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R11.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R12.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR13.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R13.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R14.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR15.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R15.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R16.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R17.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R18.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR19.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R19.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R20.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR21.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R21.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R22.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR23.BackColor = RGB(79, 98, 40)
        frmSNInfoPage.ToggleButton1R23.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R24.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R25.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R26.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R27.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R28.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R29.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R30.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R31.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1R32.BackColor = RGB(146, 208, 80)
        frmSNInfoPage.ToggleButton1aR33.BackColor = RGB(0, 176, 80)
        frmSNInfoPage.ToggleButton1R33.BackColor = RGB(146, 208, 80)
            'Orange
        frmSNInfoPage.ToggleButton2R1.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R2.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R3.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R4.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R5.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R6.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R7.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R8.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R9.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R10.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R11.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R12.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R13.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R14.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R15.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R16.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R17.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R18.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R19.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R20.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R21.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R22.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R23.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R24.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R25.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R26.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R27.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R28.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R29.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R30.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R31.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R32.BackColor = RGB(255, 192, 0)
        frmSNInfoPage.ToggleButton2R33.BackColor = RGB(255, 192, 0)
            'Blue
        frmSNInfoPage.ToggleButton3R1.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R2.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R3.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R4.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R5.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R6.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R7.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R8.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R9.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R10.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R11.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R12.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R13.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R14.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R15.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R16.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R17.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R18.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R19.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R20.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R21.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R22.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R23.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R24.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R25.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R26.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R27.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R28.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R29.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R30.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R31.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R32.BackColor = RGB(146, 205, 220)
        frmSNInfoPage.ToggleButton3R33.BackColor = RGB(146, 205, 220)

End Sub

Public Sub SNIPCurrentInfo()

    '======================================================================'
    'Copy column to backup array
    Dim bkup As Range
    For Each bkup In Worksheets("NEO 5322121").Range(Cells(1, snTitleCell.Column), Cells(56, snTitleCell.Column))
        '==========================='
        'initialize boolFlagCell and boolFlagCellClear
        boolFlagCell = False
        boolFlagCellClear = False
        'search for errors
        If (IsError(bkup.Value)) Then
            'Load SN_DeleteError form
            Set frmSNDeleteErrorSNIPCI = New SN_DeleteError
            boolCanceled = False
            'initialize txtboxes
            frmSNDeleteErrorSNIPCI.TextBox1.Value = bkup.Address(False, False)
            'define error cell range and go to error cell
            rngErrorCell = bkup.Address
            Application.Goto Worksheets("NEO 5322121").Range(rngErrorCell)
            'show form
            frmSNDeleteErrorSNIPCI.Show
            'if SN_DeleteError is exitted
            If boolCanceled Then
                Exit Sub
            End If
        End If
        'if error cell flagged
        If boolFlagCell Then: GoTo lineErrorFound
        '==========================='
        'assign value to backup array
        arraySNBackupVal(bkup.Row) = bkup.Value
        'encountering error
        If False Then
lineErrorFound:
            arraySNBackupVal(bkup.Row) = "!!!FLAGGED AS ERROR!!!"
            'reset booleans
            boolFlagCell = False
            boolFlagCellClear = False
        End If
        arraySNBackupClr(bkup.Row) = bkup.Interior.Color
    Next bkup
    '======================================================================'

    Dim cell As Range
    
    'grab serial number text
    snSearchTxt = snTitleCell.Value
    
    'Searial Number Title
    frmSNInfoPage.Caption = snSearchTxt & ":"
    
    'Serial Number Column
    snSearchCol = snTitleCell.Column
    
    'Engine Set Tracking SN Button
    If snTitleCell.Interior.Color = RGB(255, 255, 0) Then
        frmSNInfoPage.ESTButton.Value = True
    End If
    
    'Slow Moving SN Button
    If snTitleCell.Interior.Color = RGB(244, 158, 228) Then
        frmSNInfoPage.SlowButton.Value = True
    End If
    
    'row of previously completed op
    For Each cell In Worksheets("NEO 5322121").Range(Cells(7, snSearchCol), Cells(43, snSearchCol))
        If Not cell.Interior.Color = clrBlank Then
            SNIPSearchCurrentRow = cell.Row
            Exit For
        End If
    Next cell
    
    'engine set number
    frmSNInfoPage.EngSetTxtBox.Value = Worksheets("NEO 5322121").Cells(1, snSearchCol).Value
    frmSNInfoPage.EngSetTxtBox.Locked = True
    
    'engine set count
    frmSNInfoPage.EngSetCntTxtBox.Value = Worksheets("NEO 5322121").Cells(5, snSearchCol).Value
    frmSNInfoPage.EngSetCntTxtBox.Locked = True
    
    'last known location
    frmSNInfoPage.LocationTxtBox.Value = Worksheets("NEO 5322121").Cells(53, snSearchCol).Value
    
    'last op completed
    frmSNInfoPage.LastOpTxtBox.Value = Worksheets("NEO 5322121").Cells(51, snSearchCol).Value
    
    'last date seen
    SNIPLastDate = Worksheets("NEO 5322121").Cells(52, snSearchCol).Value
        'check for letters
    Dim boolSNIPLD As Boolean
    Dim j As Double
    boolSNIPLD = True
    For j = 1 To Len(SNIPLastDate)
        If Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(SNIPLastDate, j, 1)) = 0) Then
            boolSNIPLD = False
            Exit For
        End If
    Next j
        'not empty and correct format
    If Not (SNIPLastDate = "") And (Len(SNIPLastDate) > 5) And (boolSNIPLD) Then
        frmSNInfoPage.LastDateTxtBox.Value = Left(SNIPLastDate, Len(SNIPLastDate) - 5)
        'empty or incorrect format
    Else
        frmSNInfoPage.LastDateTxtBox.Value = SNIPLastDate
    End If
    
    'default info page tab and corresponding SN Column cell
    If (SNIPSearchCurrentRow >= 7) And (SNIPSearchCurrentRow <= 22) Then
        frmSNInfoPage.MultiPage1.Value = 0
        Application.Goto Worksheets("NEO 5322121").Cells(7, snSearchCol), Scroll:=True
    ElseIf ((SNIPSearchCurrentRow >= 23) And (SNIPSearchCurrentRow <= 33)) Or ((SNIPSearchCurrentRow >= 38) And (SNIPSearchCurrentRow <= 44)) Then
        frmSNInfoPage.MultiPage1.Value = 1
        Application.Goto Worksheets("NEO 5322121").Cells(23, snSearchCol), Scroll:=True
    End If

End Sub

Public Sub SNIPTxtBoxGreen()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(146, 208, 80)
        'Text Box default date and (ONLY IF ERROR WAS NOT CLEARED)
        If (txtBoxTxt = "") And Not (boolFlagCellClear) Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub SNIPTxtBoxGreenDark()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(79, 98, 40)
        'Text Box default date
        If txtBoxTxt = "" Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub SNIPTxtBoxGreenLight()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(196, 215, 155)
        'Text Box default date
        If txtBoxTxt = "" Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub SNIPTxtBoxGreenBright()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(0, 176, 80)
        'Text Box default date
        If txtBoxTxt = "" Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub SNIPTxtBoxOrange()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(255, 192, 0)
        'Text Box default date
        If txtBoxTxt = "" Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub SNIPTxtBoxBlue()

    'Colored Text Box
    If boolTog Then
        txtBoxClr = RGB(146, 205, 220)
        'Text Box default date
        If txtBoxTxt = "" Then
            txtBoxTxt = Left(TodaysDate, Len(TodaysDate) - 5)
        End If
    ElseIf Not boolTog Then
        'Uncolored Text Box
        txtBoxClr = RGB(255, 255, 255)
    End If

End Sub

Public Sub WaterFallSN()

    Dim wtrfllArClr As Long
    Dim wtrfllArVal As String
    Dim redCellSearch As Range
    Dim colRedCell As Double
    Dim arInd As Double
    Dim wtrfllTrackerRow As Double
    Dim wtrfllRange As Range
    Dim wtrfllSN As Range
    Dim colCut As Double
    Dim colInsert As Double
    
    Dim aftwtrRng As Range
    Dim aftwtrCll As Range
    Dim aftwtrSN As String
    
    'turn off screen updates
    Application.ScreenUpdating = False
    
    'search for red cell
    For Each redCellSearch In Worksheets("NEO 5322121").Range("1:1")
        'red line found
        If Worksheets("NEO 5322121").Columns(redCellSearch.Column).Interior.Color = RGB(255, 0, 0) Then
            colRedCell = redCellSearch.Column
            Exit For
        End If
    Next redCellSearch
    
    'show Waterfalling Tracker Message
    If Not (boolWtrFllDTG) Or boolWtrFllDTGFirstRun Then
        Set frmMsgWaterfall = New MsgWaterfall
    End If
    
    'message for waterfalling entire tracker
    Dim serialsPERpixel As Double
    Dim currentNumSN As Double
    Dim loadingCount As Double
    serialsPERpixel = 0
    currentNumSN = 0
    loadingCount = 0
    If boolWtrFllDTG Then
        'adjust ETA
        If secs > 0 Then
            secs = (secs - 5)
        ElseIf (secs = 0) And (mins > 0) Then
            mins = (mins - 1)
            secs = 55
        ElseIf (secs = 0) And (mins = 0) And (hrs > 0) Then
            hrs = (hrs - 1)
            mins = 59
            secs = 55
        ElseIf (secs = 0) And (mins = 0) And (hrs = 0) Then
            hrs = 0
            mins = 0
            secs = 0
        End If
        'fill in ETA Label
        If hrs > 0 Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(hrs) & " hrs, " & Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins > 0) Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins = 0) Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(secs) & " secs"
        End If
        'update loading bar
        serialsPERpixel = snCount / 166
        currentNumSN = (snCount - ((secs + (mins * 60) + (hrs * 3600)) / 5))
        loadingCount = Int(currentNumSN / serialsPERpixel)
        'hide all loading bar images
        Dim img As Double
        For img = 0 To 166
            frmMsgWaterfall.Controls("Image" & img).Visible = False
        Next img
        'unhide current loading bar image
        frmMsgWaterfall.Controls("Image" & loadingCount).Visible = True
        'change form caption
        frmMsgWaterfall.Caption = "Waterfalling Tracker..."
    End If
    
    'message if only waterfalling a group of SN's
    If Not boolWtrFllDTG Then
        frmMsgWaterfall.Label1 = "Waterfalling " & arrayWaterfallVal(6) & "..."
        frmMsgWaterfall.Image0.Visible = True
    End If
    
    'show message form
    frmMsgWaterfall.Show vbModeless
    DoEvents
    
    'search waterfall SN color array for top color
    For arInd = 13 To 43 '******Ignoring anything above EQN Entry as per Todd's request******
        'if not blank color, or a hidden cell row
        If Not (arrayWaterfallClr(arInd) = clrBlank) And ((arInd <= 33) Or (arInd >= 38)) Then
            wtrfllTrackerRow = arInd
            Exit For
        End If
    Next arInd
    
    'Check for R2O update errors
    For arInd = 7 To 43
        'if column has a blue cell
        If arrayWaterfallClr(arInd) = RGB(146, 205, 220) Then
            'update array R2O cells
            arrayWaterfallClr(49) = RGB(146, 205, 220)
            arrayWaterfallClr(50) = RGB(146, 205, 220)
            arrayWaterfallVal(49) = "R2O"
            arrayWaterfallVal(50) = arrayWaterfallVal(arInd)
            'update column R2O cells
            Worksheets("NEO 5322121").Cells(49, Int(arrayWaterfallVal(57))).Interior.Color = RGB(146, 205, 220)
            Worksheets("NEO 5322121").Cells(50, Int(arrayWaterfallVal(57))).Interior.Color = RGB(146, 205, 220)
            Worksheets("NEO 5322121").Cells(49, Int(arrayWaterfallVal(57))).Value = "R2O"
            Worksheets("NEO 5322121").Cells(50, Int(arrayWaterfallVal(57))).Value = arrayWaterfallVal(arInd)
            Exit For
        'no blue cells found
        Else
            'clear R2O cells
            arrayWaterfallClr(49) = clrBlank
            arrayWaterfallClr(50) = clrBlank
            arrayWaterfallVal(49) = ""
            arrayWaterfallVal(50) = ""
            'update column R2O cells
            Worksheets("NEO 5322121").Cells(49, Int(arrayWaterfallVal(57))).Interior.Color = clrBlank
            Worksheets("NEO 5322121").Cells(50, Int(arrayWaterfallVal(57))).Interior.Color = clrBlank
            Worksheets("NEO 5322121").Cells(49, Int(arrayWaterfallVal(57))).Value = ""
            Worksheets("NEO 5322121").Cells(50, Int(arrayWaterfallVal(57))).Value = ""
        End If
    Next arInd

    'define waterfall variables
    Set wtrfllRange = Worksheets("NEO 5322121").Range(Cells(wtrfllTrackerRow, 3), Cells(wtrfllTrackerRow, colRedCell))
    wtrfllArClr = arrayWaterfallClr(wtrfllTrackerRow)
    If arrayWaterfallVal(wtrfllTrackerRow) = "!!!FLAGGED AS ERROR!!!" Then: wtrfllArVal = ""
    If Not (arrayWaterfallVal(wtrfllTrackerRow) = "!!!FLAGGED AS ERROR!!!") Then: wtrfllArVal = arrayWaterfallVal(wtrfllTrackerRow)
    
    'iterate waterfall search row
    For Each wtrfllSN In wtrfllRange
        
        
        'if SN's top update color is green, blue RTO, purple in-house, or bright green (left)
        If (wtrfllSN.Column > 2) And ((wtrfllArClr = RGB(146, 208, 80)) Or (wtrfllArClr = RGB(146, 205, 220)) Or (wtrfllArClr = RGB(177, 160, 199)) Or (wtrfllArClr = RGB(0, 176, 80))) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(177, 160, 199)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80))) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("NEO 5322121").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(177, 160, 199)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("NEO 5322121").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        'if SN's top update color is dark green, or light green (middle)
        ElseIf (wtrfllSN.Column > 2) And ((wtrfllArClr = RGB(79, 98, 40)) Or (wtrfllArClr = RGB(196, 215, 155))) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And ((wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155))) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("NEO 5322121").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80)) Or (wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155)) Or (wtrfllSN.Interior.Color = RGB(177, 160, 199))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("NEO 5322121").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        'if SN's top update color is orange (right)
        ElseIf (wtrfllSN.Column > 2) And (wtrfllArClr = RGB(255, 192, 0)) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And (wtrfllSN.Interior.Color = RGB(255, 192, 0)) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("NEO 5322121").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("NEO 5322121").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("NEO 5322121").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80)) Or (wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155)) Or (wtrfllSN.Interior.Color = RGB(255, 192, 0)) Or (wtrfllSN.Interior.Color = RGB(177, 160, 199))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("NEO 5322121").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("NEO 5322121").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Cut Worksheets("NEO 5322121").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("NEO 5322121").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        End If
    Next wtrfllSN
    
    
lineHideWTMsg:
    'close Waterfalling Tracker Message
    If Not boolDoNotClose Then
        frmMsgWaterfall.Hide
        Unload frmMsgWaterfall
    End If
    Exit Sub
    
    
lineFindWaterfalledSNAfterwards:
    
    'set aftwtrRng
    Set aftwtrRng = Worksheets("NEO 5322121").Range("6:6")
    
    'search for SN
    For Each aftwtrCll In aftwtrRng
    
        'found SN
        If aftwtrCll.Value = aftwtrSN Then
            Application.Goto Worksheets("NEO 5322121").Cells(6, aftwtrCll.Column), Scroll:=True
            Exit For
        End If
    
    Next aftwtrCll
    
    'turn on screen updates
    Application.ScreenUpdating = True
    
    'fix white cell dates and colors
    Dim ff As Double
    Dim f As Range
    Dim ldTime As Double
    For ff = 43 To 7 Step -1
        Set f = Worksheets("NEO 5322121").Cells(ff, ActiveCell.Column)
        If Not (f.Interior.Color = RGB(146, 208, 80)) And Not (f.Interior.Color = RGB(79, 98, 40)) And Not (f.Interior.Color = RGB(196, 215, 155)) And Not (f.Interior.Color = RGB(0, 176, 80)) And Not (f.Interior.Color = RGB(255, 192, 0)) And Not (f.Interior.Color = RGB(146, 205, 220)) And Not (f.Interior.Color = RGB(177, 160, 199)) And Not (f.Interior.Color = RGB(255, 0, 0)) And Not (f.Interior.Color = RGB(0, 0, 0)) Then
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
    
    'send to hide message box and exit sub
    GoTo lineHideWTMsg
    

End Sub

Public Sub WaterFallQC()

    Dim wtrfllArClr As Long
    Dim wtrfllArVal As String
    Dim redCellSearch As Range
    Dim colRedCell As Double
    Dim arInd As Double
    Dim wtrfllTrackerRow As Double
    Dim wtrfllRange As Range
    Dim wtrfllSN As Range
    Dim colCut As Double
    Dim colInsert As Double
    
    Dim aftwtrRng As Range
    Dim aftwtrCll As Range
    Dim aftwtrSN As String
    
    'turn off screen updates
    Application.ScreenUpdating = False
    
'    'search for termination cell (called colRedCell)
'    For Each redCellSearch In Worksheets("Quality Clinic").Range("1:1")
'        'red line found
'        If Worksheets("NEO 5322121").Columns(redCellSearch.Column).Interior.Color = RGB(255, 0, 0) Then
'            colRedCell = redCellSearch.Column
'            Exit For
'        End If
'    Next redCellSearch
    
    'show Waterfalling Tracker Message
    If Not (boolWtrFllDTG) Or boolWtrFllDTGFirstRun Then
        Set frmMsgWaterfall = New MsgWaterfall
    End If
    
    'message for waterfalling entire tracker
    Dim serialsPERpixel As Double
    Dim currentNumSN As Double
    Dim loadingCount As Double
    serialsPERpixel = 0
    currentNumSN = 0
    loadingCount = 0
    If boolWtrFllDTG Then
        'adjust ETA
        If secs > 0 Then
            secs = (secs - 5)
        ElseIf (secs = 0) And (mins > 0) Then
            mins = (mins - 1)
            secs = 55
        ElseIf (secs = 0) And (mins = 0) And (hrs > 0) Then
            hrs = (hrs - 1)
            mins = 59
            secs = 55
        ElseIf (secs = 0) And (mins = 0) And (hrs = 0) Then
            hrs = 0
            mins = 0
            secs = 0
        End If
        'fill in ETA Label
        If hrs > 0 Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(hrs) & " hrs, " & Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins > 0) Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins = 0) Then
            frmMsgWaterfall.Label1.Caption = "Time Remaining:" & vbNewLine & Str(secs) & " secs"
        End If
        'update loading bar
        serialsPERpixel = snCount / 166
        currentNumSN = (snCount - ((secs + (mins * 60) + (hrs * 3600)) / 5))
        loadingCount = CDbl(currentNumSN / serialsPERpixel)
        'hide all loading bar images
        Dim img As Double
        For img = 0 To 166
            frmMsgWaterfall.Controls("Image" & img).Visible = False
        Next img
        'unhide current loading bar image
        frmMsgWaterfall.Controls("Image" & loadingCount).Visible = True
        'change form caption
        frmMsgWaterfall.Caption = "Waterfalling Tracker..."
    End If
    
    'message if only waterfalling a group of SN's
    If Not boolWtrFllDTG Then
        frmMsgWaterfall.Label1 = "Waterfalling " & arrayWaterfallVal(6) & "..."
        frmMsgWaterfall.Image0.Visible = True
    End If
    
    'show message form
    frmMsgWaterfall.Show vbModeless
    DoEvents
    
    'search waterfall SN color array for top color
    For arInd = 13 To 43 '******Ignoring anything above EQN Entry as per Todd's request******
        'if not blank color, or a hidden cell row
        If Not (arrayWaterfallClr(arInd) = clrBlank) And ((arInd <= 33) Or (arInd >= 38)) Then
            wtrfllTrackerRow = arInd
            Exit For
        End If
    Next arInd

    'define waterfall variables
    Set wtrfllRange = Worksheets("Quality Clinic").Range(Cells(wtrfllTrackerRow, 3), Cells(wtrfllTrackerRow, colRedCell))
    wtrfllArClr = arrayWaterfallClr(wtrfllTrackerRow)
    If arrayWaterfallVal(wtrfllTrackerRow) = "!!!FLAGGED AS ERROR!!!" Then: wtrfllArVal = ""
    If Not (arrayWaterfallVal(wtrfllTrackerRow) = "!!!FLAGGED AS ERROR!!!") Then: wtrfllArVal = arrayWaterfallVal(wtrfllTrackerRow)
    
    'iterate waterfall search row
    For Each wtrfllSN In wtrfllRange
        
        
        'if SN's top update color is green, blue RTO, or bright green (left)
        If (wtrfllSN.Column > 2) And ((wtrfllArClr = RGB(146, 208, 80)) Or (wtrfllArClr = RGB(146, 205, 220)) Or (wtrfllArClr = RGB(0, 176, 80))) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80))) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("Quality Clinic").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("Quality Clinic").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        'if SN's top update color is dark green, or light green (middle)
        ElseIf (wtrfllSN.Column > 2) And ((wtrfllArClr = RGB(79, 98, 40)) Or (wtrfllArClr = RGB(196, 215, 155))) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And ((wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155))) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("Quality Clinic").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80)) Or (wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("Quality Clinic").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        'if SN's top update color is orange (right)
        ElseIf (wtrfllSN.Column > 2) And (wtrfllArClr = RGB(255, 192, 0)) Then
            'SN placement found (not at end of section)
            If ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not (IsEmpty(wtrfllSN)) And (wtrfllSN.Interior.Color = RGB(255, 192, 0)) Then
                'found first later dated group color
                If DateValue(wtrfllSN.Value) > DateValue(wtrfllArVal) Then
                    'define cut and insert column numbers
                    colCut = Int(arrayWaterfallVal(57))
                    colInsert = wtrfllSN.Column
                    If (colInsert <= colCut) Then: colCut = (colCut + 1)
                    'cut and insert SN column=================================================================
                    aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                    Worksheets("Quality Clinic").Columns(colInsert).Insert
                    'loading bar update 1 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image0.Visible = False
                        frmMsgWaterfall.Image56.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                    'loading bar update 2 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image56.Visible = False
                        frmMsgWaterfall.Image111.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    Worksheets("Quality Clinic").Columns(colCut).Delete
                    'loading bar update 3 (only for non entire tracker waterfalling)
                    If Not boolWtrFllDTG Then
                        frmMsgWaterfall.Image111.Visible = False
                        frmMsgWaterfall.Image166.Visible = True
                        frmMsgWaterfall.Show vbModeless
                        DoEvents
                    End If
                    'cut and insert SN column=================================================================
                    GoTo lineFindWaterfalledSNAfterwards
                End If
            'SN placement found at end of section (found first non-group color)
            ElseIf ((wtrfllSN.Row = 13) Or (Worksheets("Quality Clinic").Range(Cells(13, wtrfllSN.Column), wtrfllSN.Offset(-1, 0)).Interior.Color = clrBlank)) And Not ((wtrfllSN.Interior.Color = RGB(146, 208, 80)) Or (wtrfllSN.Interior.Color = RGB(146, 205, 220)) Or (wtrfllSN.Interior.Color = RGB(0, 176, 80)) Or (wtrfllSN.Interior.Color = RGB(79, 98, 40)) Or (wtrfllSN.Interior.Color = RGB(196, 215, 155)) Or (wtrfllSN.Interior.Color = RGB(255, 192, 0))) Then
                'define cut and insert column numbers
                colCut = Int(arrayWaterfallVal(57))
                colInsert = wtrfllSN.Column
                If (colInsert <= colCut) Then: colCut = (colCut + 1)
                'cut and insert SN column=================================================================
                aftwtrSN = Worksheets("Quality Clinic").Cells(6, Int(arrayWaterfallVal(57))).Value
                Worksheets("Quality Clinic").Columns(colInsert).Insert
                'loading bar update 1 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image0.Visible = False
                    frmMsgWaterfall.Image56.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Cut Worksheets("Quality Clinic").Columns(colInsert)
                'loading bar update 2 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image56.Visible = False
                    frmMsgWaterfall.Image111.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                Worksheets("Quality Clinic").Columns(colCut).Delete
                'loading bar update 3 (only for non entire tracker waterfalling)
                If Not boolWtrFllDTG Then
                    frmMsgWaterfall.Image111.Visible = False
                    frmMsgWaterfall.Image166.Visible = True
                    frmMsgWaterfall.Show vbModeless
                    DoEvents
                End If
                'cut and insert SN column=================================================================
                GoTo lineFindWaterfalledSNAfterwards
            End If
        
        
        End If
    Next wtrfllSN
    
    
lineHideWTMsg:
    'close Waterfalling Tracker Message
    If Not boolDoNotClose Then
        frmMsgWaterfall.Hide
        Unload frmMsgWaterfall
    End If
    Exit Sub
    
    
lineFindWaterfalledSNAfterwards:
    
    'set aftwtrRng
    Set aftwtrRng = Worksheets("Quality Clinic").Range("6:6")
    
    'search for SN
    For Each aftwtrCll In aftwtrRng
    
        'found SN
        If aftwtrCll.Value = aftwtrSN Then
            Application.Goto Worksheets("Quality Clinic").Cells(6, aftwtrCll.Column), Scroll:=True
            Exit For
        End If
    
    Next aftwtrCll
    
    'turn on screen updates
    Application.ScreenUpdating = True
    
    'fix white cell dates and colors
    Dim ff As Double
    Dim f As Range
    Dim ldTime As Double
    For ff = 43 To 7 Step -1
        Set f = Worksheets("NEO 5322121").Cells(ff, ActiveCell.Column)
        If Not (f.Interior.Color = RGB(146, 208, 80)) And Not (f.Interior.Color = RGB(79, 98, 40)) And Not (f.Interior.Color = RGB(196, 215, 155)) And Not (f.Interior.Color = RGB(0, 176, 80)) And Not (f.Interior.Color = RGB(255, 192, 0)) And Not (f.Interior.Color = RGB(146, 205, 220)) And Not (f.Interior.Color = RGB(177, 160, 199)) And Not (f.Interior.Color = RGB(255, 0, 0)) And Not (f.Interior.Color = RGB(0, 0, 0)) Then
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
    
    'send to hide message box and exit sub
    GoTo lineHideWTMsg


End Sub

Public Sub DeleteBlackColumns()

    Dim esCell As Range
    Dim esCellCol As Double
    Dim esRange As Range
    Dim esNumber As Long
    Dim s As Range
    Dim img As Double
    Dim SNTotal As Double
    Dim loadingCount As Double
    Dim cntBlackLines As Double
    
    'initialize variables
    Set esRange = Worksheets("NEO 5322121").Range("1:1")
    esNumber = Worksheets("NEO 5322121").Cells(1, 3)
    SNTotal = 0
    cntBlackLines = 0
    
    'count all SN's
    For Each s In Worksheets("NEO 5322121").Range("6:6")
        If s.Interior.Color = RGB(255, 0, 0) Then
            'remove 1st two columns
            SNTotal = SNTotal - 2
            Exit For
        ElseIf s.Interior.Color = RGB(0, 0, 0) Then
            cntBlackLines = cntBlackLines + 1
        End If
        SNTotal = SNTotal + 1
    Next s
    
    'exit sub if no black lines detected
    If cntBlackLines = 0 Then
        Exit Sub
    End If
    
    'show Temporarily Deleting Black Lines Message
    Set frmMsgTempDelBL = New MsgTempDelBL
    frmMsgTempDelBL.Image0.Visible = True
    
    For Each esCell In esRange
        'Do Not Iterate first 2 columns
        If esCell.Column > 2 Then
            
            
            'update loading bar
            For img = 0 To 166
                frmMsgTempDelBL.Controls("Image" & img).Visible = False
            Next img
            loadingCount = Int((esCell.Column - 2) / (SNTotal / 166))
            frmMsgTempDelBL.Controls("Image" & loadingCount).Visible = True
            frmMsgTempDelBL.Show vbModeless
            DoEvents
            
            
            'Red Column found
            If (esCell.Interior.Color = RGB(255, 0, 0)) Then
                'end of waterfall found, so exit iteration loop
                Exit For
                
                
            'if engine set black line found
            ElseIf (esCell.Interior.Color = RGB(0, 0, 0)) Then
                'delete black column
                Worksheets("NEO 5322121").Columns(esCell.Column).Delete
            
            
            End If
        End If
    Next esCell
    
    
    'close Temporarily Deleting Black Lines Message
    frmMsgTempDelBL.Hide


End Sub

Public Sub EngineSetHandler()

    Dim esCell As Range
    Dim esCellCol As Double
    Dim esRange As Range
    Dim esNumber As Double
    Dim esCount As Double
    Dim s As Range
    Dim img As Double
    Dim SNTotal As Double
    Dim loadingCount As Double
    
    
    'initialize variables
    Set esRange = Worksheets("NEO 5322121").Range("1:1")
    esCount = 1
    SNTotal = 0
    
    'count all SN's
    For Each s In Worksheets("NEO 5322121").Range("6:6")
        If s.Interior.Color = RGB(255, 0, 0) Then
            'remove 1st two columns
            SNTotal = SNTotal - 2
            Exit For
        End If
        SNTotal = SNTotal + 1
    Next s
    
    'look up next engine set
    Dim curengsetRng As Range
    Dim curengsetCell As Range
    Set curengsetRng = Worksheets("Shipped").Range("1:1")
    For Each curengsetCell In curengsetRng
        'if more current engine set found
        If IsNumeric(curengsetCell.Value) And (curengsetCell.Value > 0) Then
            esNumber = curengsetCell.Value + 1
        End If
    Next curengsetCell
    
    'show Separating Engine Sets Message
    Set frmMsgSeparatingEngSets = New MsgSeparatingEngSets
    frmMsgSeparatingEngSets.Image0.Visible = True
    
    
    For Each esCell In esRange
        'Do Not Iterate first 2 columns
        If esCell.Column > 2 Then
        
        
            'update loading bar
            For img = 0 To 166
                frmMsgSeparatingEngSets.Controls("Image" & img).Visible = False
            Next img
            loadingCount = Int((esCell.Column - 2) / (SNTotal / 166))
            frmMsgSeparatingEngSets.Controls("Image" & loadingCount).Visible = True
            frmMsgSeparatingEngSets.Show vbModeless
            DoEvents


            'Red Column found
            If (esCell.Interior.Color = RGB(255, 0, 0)) Then
                'scroll back to beginning of tracker
                Application.Goto Worksheets("NEO 5322121").Cells(1, 3), Scroll:=True
                'end of waterfall found, so exit iteration loop
                Exit For
                
line770152:
            'if engine set of 33 blades
            ElseIf (esNumber <= 770152) And Not (esCell.Interior.Color = RGB(255, 0, 0)) Then
                'go to esCell
                Application.Goto Worksheets("NEO 5322121").Cells(1, esCell.Column), Scroll:=True
                'when engine set counts complete
                If esCount = 34 Then
                    'insert black line
                    If Not esCell.Interior.Color = RGB(0, 0, 0) Then
                        Worksheets("NEO 5322121").Columns(esCell.Column).Insert
                        Worksheets("NEO 5322121").Columns(esCell.Offset(0, -1).Column).Interior.Color = RGB(0, 0, 0)
                    End If
                    'reset esCount and increase esNumber and increase SNTotal
                    esCount = 1
                    esNumber = esNumber + 1
                    SNTotal = SNTotal + 1
                    'send to next esCell
                    GoTo lineNextesCell
                End If
                'Apply Engine Set Number
                esCell.Value = esNumber
                'Apply Engine Set Count
                esCell.Offset(4, 0).Value = esCount
                'increase esCount
                esCount = esCount + 1

line770153:
            'if engine set of 30 blades
            ElseIf (esNumber >= 770153) And Not (esCell.Interior.Color = RGB(255, 0, 0)) Then
                'go to esCell
                Application.Goto Worksheets("NEO 5322121").Cells(1, esCell.Column), Scroll:=True
                'when engine set counts complete
                If esCount = 31 Then
                    'insert black line
                    If Not esCell.Interior.Color = RGB(0, 0, 0) Then
                        Worksheets("NEO 5322121").Columns(esCell.Column).Insert
                        Worksheets("NEO 5322121").Columns(esCell.Offset(0, -1).Column).Interior.Color = RGB(0, 0, 0)
                    End If
                    'reset esCount and increase esNumber and increase SNTotal
                    esCount = 1
                    esNumber = esNumber + 1
                    SNTotal = SNTotal + 1
                    'send to next esCell
                    GoTo lineNextesCell
                End If
                'Apply Engine Set Number
                esCell.Value = esNumber
                'Apply Engine Set Count
                esCell.Offset(4, 0).Value = esCount
                'increase esCount
                esCount = esCount + 1
            
            
            End If
        End If
lineNextesCell:
    Next esCell
    
    
    'close Separating Engine Sets Message
    frmMsgSeparatingEngSets.Hide
    

End Sub

