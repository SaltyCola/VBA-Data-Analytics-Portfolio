VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TC_UCDisplay 
   Caption         =   "Unit Column Display:"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   OleObjectBlob   =   "TC_UCDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TC_UCDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public UCShown As TC_UnitColumn 'Unit Column being displayed by the userform.
Public bMultiUpdate As Boolean 'boolean to tell UC Display to show applied to all message
Public bValidatedHeaders As Boolean 'True: Info in Headers Frame is valid.
Public bValidatedOperations As Boolean 'True: Info in Operations Frame is valid.
Public bValidatedNotes As Boolean 'True: Info in Notes Frame is valid.

Private Sub btnMainMenu_Click()

    Dim bContinue As Boolean 'True: Yes clicked on ufYesNoMsg ; False: No clicked on ufYesNoMsg
    
    'initialize boolean
    bContinue = True
    
    'if save changes button is untoggled ask for confirmation of choice to move on without saving
    If Me.tglSaveChanges.Value = False Then
        ufYesNoMsg.YesNoMsgInitialize ("Any unsaved Unit Column changes will be lost." & vbNewLine & "Continue?")
        'No clicked, cancel next if clause
        If Not ufYesNoMsg.bYesNoMsg Then
            bContinue = False
        End If
    End If
    
    'continue with code (do not allow movement of tracker while main menu is shown)
    If bContinue Then
        'Data Upload to ArrWIP
        If Me.tglSaveChanges.Value = True Then
            Me.ApplyUCChanges
        End If
        'Multi-Update Message
        If Me.tglSaveChanges.Value = True And Me.bMultiUpdate Then
            MsgBox "Changes will be applied to " & UBound(ArrListBox) & " Unit Columns after waterfalling."
        End If
        'Reset UC Display and corresponding variables
        Erase ArrListBox() 'clear ArrListBox
        Me.bMultiUpdate = False 'reset Multi-Update boolean
        Me.ClearUCData 'clear ufUCDisplay for later use
        Me.StopScrollHandler 'Cut mousewheel scroll tie
        Me.Hide 'hide UC Display
        ufMainMenu.Show 'show Main Menu (without tracker movement)
    End If

End Sub

Private Sub btnSearchNew_Click()

    Dim bContinue As Boolean 'True: Yes clicked on ufYesNoMsg ; False: No clicked on ufYesNoMsg
    
    'initialize boolean
    bContinue = True
    
    'if save changes button is untoggled ask for confirmation of choice to move on without saving
    If Me.tglSaveChanges.Value = False Then
        ufYesNoMsg.YesNoMsgInitialize ("Any unsaved Unit Column changes will be lost. Continue?")
        'No clicked, cancel next if clause
        If Not ufYesNoMsg.bYesNoMsg Then
            bContinue = False
        End If
    End If
    
    'continue with code
    If bContinue Then
        'Data Upload to ArrWIP
        If Me.tglSaveChanges.Value = True Then
            Me.ApplyUCChanges
        End If
        'Multi-Update Message
        If Me.tglSaveChanges.Value = True And Me.bMultiUpdate Then
            MsgBox "Changes will be applied to " & UBound(ArrListBox) & " Unit Columns after waterfalling."
        End If
        'Reset UC Display and corresponding variables
        Erase ArrListBox() 'clear ArrListBox
        Me.bMultiUpdate = False 'reset Multi-Update boolean
        Me.ClearUCData 'clear ufUCDisplay for later use
        Me.StopScrollHandler 'Cut mousewheel scroll tie
        Me.Hide 'hide UC Display
        ufUCSearch.Show 'Show UC Search (without tracker movement)
    End If

End Sub

Public Sub ApplyUCChanges() '<==Alter (Entire Sub)
'Sub Fired only when leaving UC Display with "Save Changes" Toggle Button clicked.
'Check for altered data, add to corresponding changes array, then update corresponding UC/s.
'For Dates: Changes currently receive the year of today's date.
'For Dates: Test for Changes using cstr() of the values and IGNORE "/YEAR"
'For Last Date Seen: If no change to last date seen textbox, then apply today's date.


    Dim arrUCChanges_Headers_Values() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges_Headers_Colors() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges_Operations_Values() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges_Operations_Colors() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges_Notes_Values() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges_Notes_Colors() As Variant 'Changes made to current UCShown in ufUCDisplay
    Dim arrUCChanges(1 To 3, 1 To 2) As Variant 'Matrix Array holding all above arrays
    
    Dim cIndex As Integer 'Index of current textbox control (corresponding index to UCShown)
    Dim ucIndex As Integer 'Index of UC to update in ArrWIP
    
    Dim strValidDate As String 'String variable to send through Date Validation Function before assigning to changes array
    
    Dim bChangesMade As Boolean 'True: A Change was made for the Displayed UC
    
    Dim ctrl As Control 'iterator
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    
    
    'Reset Changes Arrays
    ReDim arrUCChanges_Headers_Values(1 To Me.UCShown.Headers.GroupSize, 1 To 2) As Variant 'Strings and Booleans
    ReDim arrUCChanges_Headers_Colors(1 To Me.UCShown.Headers.GroupSize, 1 To 2) As Variant 'Longs and Booleans
    ReDim arrUCChanges_Operations_Values(1 To Me.UCShown.NumberOfOps, 1 To 2) As Variant 'Strings and Booleans
    ReDim arrUCChanges_Operations_Colors(1 To Me.UCShown.NumberOfOps, 1 To 2) As Variant 'Longs and Booleans
    ReDim arrUCChanges_Notes_Values(1 To Me.UCShown.Notes.GroupSize, 1 To 2) As Variant 'Strings and Booleans
    ReDim arrUCChanges_Notes_Colors(1 To Me.UCShown.Notes.GroupSize, 1 To 2) As Variant 'Longs and Booleans
    
    
    'place into Changes Array Matrix
    arrUCChanges(1, 1) = arrUCChanges_Headers_Values
    arrUCChanges(1, 2) = arrUCChanges_Headers_Colors
    arrUCChanges(2, 1) = arrUCChanges_Operations_Values
    arrUCChanges(2, 2) = arrUCChanges_Operations_Colors
    arrUCChanges(3, 1) = arrUCChanges_Notes_Values
    arrUCChanges(3, 2) = arrUCChanges_Notes_Colors
    
    
    'initialize all booleans to false (no change)
    bChangesMade = False
    For i1 = 1 To UBound(arrUCChanges)
        For i2 = 1 To UBound(arrUCChanges(i1, 1))
            arrUCChanges(i1, 1)(i2, 2) = False
            arrUCChanges(i1, 2)(i2, 2) = False
        Next i2
    Next i1
    
    
    
'========= check userform for altered data and store in corresponding changes arrays =========
        
    'Headers Changes
    For Each ctrl In Me.frameHeaders.Controls
        'textboxes only
        If TypeName(ctrl) = "TextBox" Then
            'grab ctrl index
            cIndex = Int(Replace(ctrl.Name, "txtH", ""))
            'value change?
            If CStr(ctrl.Value) <> CStr(Me.UCShown.Headers.ValuesList(cIndex)) Then
                'boolean
                arrUCChanges(1, 1)(cIndex, 2) = True
                'value
                arrUCChanges(1, 1)(cIndex, 1) = CStr(ctrl.Value)
            End If
            'color change?
            If CLng(ctrl.BackColor) <> CLng(Me.UCShown.Headers.ColorsList(cIndex)) Then
                'boolean
                arrUCChanges(1, 2)(cIndex, 2) = True
                'value
                arrUCChanges(1, 2)(cIndex, 1) = CLng(ctrl.BackColor)
            End If
        End If
    Next ctrl
    
    'Operations Changes
    For Each ctrl In Me.frameOperations.Controls
        'textboxes only
        If TypeName(ctrl) = "TextBox" Then
            'grab ctrl index
            cIndex = Int(Replace(ctrl.Name, "txtOps", ""))
            'value change? {subtract year from UCDate with: left( ..., (len(...) - 5) )}
                'UCDate <> 0-date, and there is a change
                If (Me.UCShown.OperationsList(cIndex).UCDate <> CDate(0)) And (CStr(ctrl.Value) <> Left(CStr(Me.UCShown.OperationsList(cIndex).UCDate), (Len(CStr(Me.UCShown.OperationsList(cIndex).UCDate)) - 5))) Then
                    'boolean
                    arrUCChanges(2, 1)(cIndex, 2) = True
                    'value
                        'value in textbox was deleted (still a change)
                        If CStr(ctrl.Value) = "" Then
                            arrUCChanges(2, 1)(cIndex, 1) = CStr(ctrl.Value)
                        'add year to date string (min len w/o yr: 3 ; max len w/o yr: 5)
                        ElseIf Len(CStr(ctrl.Value)) >= 3 And Len(CStr(ctrl.Value)) <= 5 And IsDate(CStr(ctrl.Value)) Then
                            arrUCChanges(2, 1)(cIndex, 1) = CStr(ctrl.Value) & Right(CStr(Date), 5)
                        'don't add year to date string (min len w/ yr: 8 ; max len w/ yr: 10)
                        ElseIf Len(CStr(ctrl.Value)) >= 8 And Len(CStr(ctrl.Value)) <= 10 And IsDate(CStr(ctrl.Value)) Then
                            arrUCChanges(2, 1)(cIndex, 1) = CStr(ctrl.Value)
                        'reverse boolean's value to no change
                        Else
                            arrUCChanges(2, 1)(cIndex, 2) = False
                        End If
                'UCDate = 0-date, and txtbox.value <> ""
                ElseIf (Me.UCShown.OperationsList(cIndex).UCDate = CDate(0)) And (CStr(ctrl.Value) <> "") Then
                    'boolean
                    arrUCChanges(2, 1)(cIndex, 2) = True
                    'value
                        'add year to date string (min len w/o yr: 3 ; max len w/o yr: 5)
                        If Len(CStr(ctrl.Value)) >= 3 And Len(CStr(ctrl.Value)) <= 5 And IsDate(CStr(ctrl.Value)) Then
                            arrUCChanges(2, 1)(cIndex, 1) = CStr(ctrl.Value) & Right(CStr(Date), 5)
                        'don't add year to date string (min len w/ yr: 8 ; max len w/ yr: 10)
                        ElseIf Len(CStr(ctrl.Value)) >= 8 And Len(CStr(ctrl.Value)) <= 10 And IsDate(CStr(ctrl.Value)) Then
                            arrUCChanges(2, 1)(cIndex, 1) = CStr(ctrl.Value)
                        'reverse boolean's value to no change
                        Else
                            arrUCChanges(2, 1)(cIndex, 2) = False
                        End If
                End If
            'color change?
            If CLng(ctrl.BackColor) <> CLng(Me.UCShown.OperationsList(cIndex).UCColor) Then
                'boolean
                arrUCChanges(2, 2)(cIndex, 2) = True
                'value
                arrUCChanges(2, 2)(cIndex, 1) = CLng(ctrl.BackColor)
            End If
        End If
    Next ctrl
    
    'Notes Changes
    For Each ctrl In Me.frameNotes.Controls
        'textboxes only
        If TypeName(ctrl) = "TextBox" Then
            'grab ctrl index
            cIndex = Int(Replace(ctrl.Name, "txtN", ""))
            'value change?
            If CStr(ctrl.Value) <> CStr(Me.UCShown.Notes.ValuesList(cIndex)) Then
                'boolean
                arrUCChanges(3, 1)(cIndex, 2) = True
                'value
                arrUCChanges(3, 1)(cIndex, 1) = CStr(ctrl.Value)
            End If
            'color change?
            If CLng(ctrl.BackColor) <> CLng(Me.UCShown.Notes.ColorsList(cIndex)) Then
                'boolean
                arrUCChanges(3, 2)(cIndex, 2) = True
                'value
                arrUCChanges(3, 2)(cIndex, 1) = CLng(ctrl.BackColor)
            End If
        End If
    Next ctrl
        
'=============================================================================================



    'Check for any changes at all
    For i1 = 1 To UBound(arrUCChanges)
        For i2 = 1 To UBound(arrUCChanges(i1, 1))
            If (arrUCChanges(i1, 1)(i2, 2) = True) Or (arrUCChanges(i1, 2)(i2, 2) = True) Then
                'changes were made to this UC
                bChangesMade = True
                Exit For 'only need one change to alter bChangesMade
            End If
        Next i2
    Next i1



'======================== Apply Changes to UC/s in ArrWIP ========================


    'Single Update =======================================================
    If (bChangesMade) And (Me.bMultiUpdate = False) Then
        
        'Search for UC to change
        For i1 = 1 To UBound(ArrWIP)
            'UC found
            If ArrWIP(i1).TrackingNumber = Me.UCShown.TrackingNumber Then
                ucIndex = i1
                Exit For
            End If
        Next i1
        
        'Last Date Seen = Today's Date if it was not altered by user
        If arrUCChanges(3, 1)(4, 2) = False Then 'no user-changes to last date seen
            ArrWIP(ucIndex).LastDateSeen = Date 'alter last date seen
            arrUCChanges(3, 1)(4, 1) = CStr(Date) 'alter in changes array to give to notes.valueslist as well
            arrUCChanges(3, 1)(4, 2) = True 'alter corresponding changes array boolean to true
        Else 'apply user-changed last date seen
            If IsDate(arrUCChanges(3, 1)(4, 1)) Then
                ArrWIP(ucIndex).LastDateSeen = CDate(arrUCChanges(3, 1)(4, 1)) 'update last date seen prop with change made
            Else
                ArrWIP(ucIndex).LastDateSeen = CDate(0)
            End If
        End If
        
        'Apply Headers Changes
        For i1 = 1 To ArrWIP(ucIndex).Headers.GroupSize
            'apply value changes
            If arrUCChanges(1, 1)(i1, 2) = True Then
                ArrWIP(ucIndex).Headers.ValuesList(i1) = arrUCChanges(1, 1)(i1, 1)
            End If
            'apply color changes
            If arrUCChanges(1, 2)(i1, 2) = True Then
                ArrWIP(ucIndex).Headers.ColorsList(i1) = arrUCChanges(1, 2)(i1, 1)
            End If
        Next i1
        
        'Apply Operations Changes
        For i1 = 1 To ArrWIP(ucIndex).NumberOfOps
            'apply value changes if it's a date
            If arrUCChanges(2, 1)(i1, 2) = True And IsDate(arrUCChanges(2, 1)(i1, 1)) Then
                ArrWIP(ucIndex).OperationsList(i1).UCDate = CDate(arrUCChanges(2, 1)(i1, 1))
            'apply value changes if it's empty, but still a change
            ElseIf arrUCChanges(2, 1)(i1, 2) = True And arrUCChanges(2, 1)(i1, 1) = "" Then
                ArrWIP(ucIndex).OperationsList(i1).UCDate = CDate(0)
            End If
            'apply color changes
            If arrUCChanges(2, 2)(i1, 2) = True Then
                ArrWIP(ucIndex).OperationsList(i1).UCColor = arrUCChanges(2, 2)(i1, 1)
            End If
        Next i1
        
        'Apply Notes Changes
        For i1 = 1 To ArrWIP(ucIndex).Notes.GroupSize
            'apply value changes
            If arrUCChanges(3, 1)(i1, 2) = True Then
                ArrWIP(ucIndex).Notes.ValuesList(i1) = arrUCChanges(3, 1)(i1, 1)
            End If
            'apply color changes
            If arrUCChanges(3, 2)(i1, 2) = True Then
                ArrWIP(ucIndex).Notes.ColorsList(i1) = arrUCChanges(3, 2)(i1, 1)
            End If
        Next i1
            
        'Apply Final UC Property Changes
        Me.InfoUpdate_LastOpCompleted (ucIndex)
        Me.InfoUpdate_ColorOrderIndex (ucIndex)
        Me.InfoUpdate_RTO (ucIndex)
    
    
    'Multi-Update =======================================================
    ElseIf (bChangesMade) And (Me.bMultiUpdate = True) Then
        'iterate list of UCs to update from ArrListBox
        For i1 = 1 To UBound(ArrListBox)
            
            'Search for UC to change
            For i2 = 1 To UBound(ArrWIP)
                'UC found
                If ArrWIP(i2).TrackingNumber = ArrListBox(i1).TrackingNumber Then
                    ucIndex = i2
                    Exit For
                End If
            Next i2
            
            'Last Date Seen = Today's Date if it was not altered by user
            If arrUCChanges(3, 1)(4, 2) = False Then 'no user-changes to last date seen
                ArrWIP(ucIndex).LastDateSeen = Date 'alter last date seen
                arrUCChanges(3, 1)(4, 1) = CStr(Date) 'alter in changes array to give to notes.valueslist as well
                arrUCChanges(3, 1)(4, 2) = True 'alter corresponding changes array boolean to true
            Else 'apply user-changed last date seen
                If IsDate(arrUCChanges(3, 1)(4, 1)) Then
                    ArrWIP(ucIndex).LastDateSeen = CDate(arrUCChanges(3, 1)(4, 1)) 'update last date seen prop with change made
                Else
                    ArrWIP(ucIndex).LastDateSeen = CDate(0)
                End If
            End If
            
            'Apply Headers Changes
            For i2 = 1 To ArrWIP(ucIndex).Headers.GroupSize
                'apply value changes
                If arrUCChanges(1, 1)(i2, 2) = True Then
                    ArrWIP(ucIndex).Headers.ValuesList(i2) = arrUCChanges(1, 1)(i2, 1)
                End If
                'apply color changes
                If arrUCChanges(1, 2)(i2, 2) = True Then
                    ArrWIP(ucIndex).Headers.ColorsList(i2) = arrUCChanges(1, 2)(i2, 1)
                End If
            Next i2
            
            'Apply Operations Changes
            For i2 = 1 To ArrWIP(ucIndex).NumberOfOps
                'apply value changes if it's a date
                If arrUCChanges(2, 1)(i2, 2) = True And IsDate(arrUCChanges(2, 1)(i2, 1)) Then
                    ArrWIP(ucIndex).OperationsList(i2).UCDate = CDate(arrUCChanges(2, 1)(i2, 1))
                'apply value changes if it's empty, but still a change
                ElseIf arrUCChanges(2, 1)(i2, 2) = True And arrUCChanges(2, 1)(i2, 1) = "" Then
                    ArrWIP(ucIndex).OperationsList(i2).UCDate = CDate(0)
                End If
                'apply color changes
                If arrUCChanges(2, 2)(i2, 2) = True Then
                    ArrWIP(ucIndex).OperationsList(i2).UCColor = arrUCChanges(2, 2)(i2, 1)
                End If
            Next i2
            
            'Apply Notes Changes
            For i2 = 1 To ArrWIP(ucIndex).Notes.GroupSize
                'apply value changes
                If arrUCChanges(3, 1)(i2, 2) = True Then
                    ArrWIP(ucIndex).Notes.ValuesList(i2) = arrUCChanges(3, 1)(i2, 1)
                End If
                'apply color changes
                If arrUCChanges(3, 2)(i2, 2) = True Then
                    ArrWIP(ucIndex).Notes.ColorsList(i2) = arrUCChanges(3, 2)(i2, 1)
                End If
            Next i2
            
            'Apply Final UC Property Changes
            Me.InfoUpdate_LastOpCompleted (ucIndex)
            Me.InfoUpdate_ColorOrderIndex (ucIndex)
            Me.InfoUpdate_RTO (ucIndex)
        
        Next i1
    
    
    End If

'=================================================================================



End Sub

Public Sub InfoUpdate_LastOpCompleted(ByVal ucIndex As Integer)

    Dim i As Integer 'iterator
    
    'Last Op Completed
    For i = 1 To ArrWIP(ucIndex).NumberOfOps
        'look for first non blank and enabled OpRow
        If ArrWIP(ucIndex).OperationsList(i).UCColor <> ArrWIP(ucIndex).OperationsList(1).UCColorList.Blank And ArrWIP(ucIndex).OperationsList(i).Enabled Then
            'last op completed
            ArrWIP(ucIndex).LastOpCompleted = i
            Exit For 'end for loop
        End If
    Next i

End Sub

Public Sub InfoUpdate_ColorOrderIndex(ByVal ucIndex As Integer)

    Dim i As Integer 'iterator
    Dim iLastOp As Integer 'index for UC's last op
    
    'set iLastOp
    iLastOp = ArrWIP(ucIndex).LastOpCompleted
    
    'Color Order Index (iterate color order array from CellColor Class)
    For i = 1 To UBound(ArrWIP(ucIndex).OperationsList(iLastOp).UCColorList.arrColorOrder, 2)
        'color option found in cell color class's color order array
        If ArrWIP(ucIndex).OperationsList(iLastOp).UCColor = ArrWIP(ucIndex).OperationsList(iLastOp).UCColorList.arrColorOrder(1, i) Then
            'apply corresponding color order index to UC
            ArrWIP(ucIndex).ColorOrderIndex = ArrWIP(ucIndex).OperationsList(iLastOp).UCColorList.arrColorOrder(2, i)
            Exit For 'end for loop
        End If
    Next i

End Sub

Public Sub InfoUpdate_RTO(ByVal ucIndex As Integer) '<==Alter (Entire Sub, Unique)

    Dim i As Integer 'iterator
    Dim clrRTOBlue As Long 'temp color variable to aleviate code space
    
    'initialize color variable
    clrRTOBlue = ArrWIP(ucIndex).OperationsList(1).UCColorList.RTO
    
    'RTO rows (in UC's Notes class)
    For i = 1 To ArrWIP(ucIndex).NumberOfOps
        'look for first RTO Blue and enabled OpRow
        If ArrWIP(ucIndex).OperationsList(i).UCColor = ArrWIP(ucIndex).OperationsList(1).UCColorList.RTO And ArrWIP(ucIndex).OperationsList(i).Enabled Then
            'update two rows in UC.Notes
            ArrWIP(ucIndex).Notes.ValuesList(1) = "R2O"
            ArrWIP(ucIndex).Notes.ColorsList(1) = clrRTOBlue
            ArrWIP(ucIndex).Notes.ValuesList(2) = ArrWIP(ucIndex).OperationsList(i).UCDate
            ArrWIP(ucIndex).Notes.ColorsList(2) = clrRTOBlue
            Exit For 'end for loop
        End If
    Next i

End Sub

Private Sub tglH_EditLock_Click()
'Lock/Unlock the Headers Frame for editting.

    'change lock/unlock picture
    If Me.tglH_EditLock.Value = True Then
        Me.imgLocked.Visible = True
        Me.imgUnlocked.Visible = False
    ElseIf Me.tglH_EditLock.Value = False Then
        Me.imgLocked.Visible = False
        Me.imgUnlocked.Visible = True
    End If
    
    'lock/unlock controls in the Headers frame
    For Each ctrl In Me.frameHeaders.Controls
        'lock only textboxes
        If TypeName(ctrl) = "TextBox" Then
            If Me.tglH_EditLock.Value = True Then: ctrl.Locked = True
            If Me.tglH_EditLock.Value = False Then: ctrl.Locked = False
        End If
    Next ctrl

End Sub

Private Sub tglSaveChanges_Click()
'Lock all controls except tglSaveChanges, btnSearchNew and btnMainMenu

    Dim ctrl As Control 'iterator
    
    'disable/enable headers lock toggle button
    If Me.tglSaveChanges.Value = True Then: Me.tglH_EditLock.Enabled = False
    If Me.tglSaveChanges.Value = False Then: Me.tglH_EditLock.Enabled = True
    
    'disable/enable controls in the Headers frame
    For Each ctrl In Me.frameHeaders.Controls
        If Me.tglSaveChanges.Value = True Then: ctrl.Enabled = False
        If Me.tglSaveChanges.Value = False Then: ctrl.Enabled = True
    Next ctrl
    
    'disable/enable controls in the Operations frame
    For Each ctrl In Me.frameOperations.Controls
        If Me.tglSaveChanges.Value = True Then: ctrl.Enabled = False
        If Me.tglSaveChanges.Value = False Then: ctrl.Enabled = True
        'always disable hidden op rows
            'op row labels
            If ctrl.BackColor = RGB(0, 0, 0) And ctrl.ForeColor = RGB(255, 255, 255) Then
                ctrl.Enabled = False
            'textboxes
            ElseIf TypeName(ctrl) = "TextBox" Then
                If Me.frameOperations.Controls("lblOps" & Replace(ctrl.Name, "txtOps", "")).Enabled = False Then: ctrl.Enabled = False
            'color buttons
            ElseIf TypeName(ctrl) = "CommandButton" Then
                If Me.frameOperations.Controls("lblOps" & Replace(ctrl.Name, "btnOps", "")).Enabled = False Then: ctrl.Enabled = False
            End If
    Next ctrl
    
    'disable/enable controls in the Notes frame
    For Each ctrl In Me.frameNotes.Controls
        If Me.tglSaveChanges.Value = True Then: ctrl.Enabled = False
        If Me.tglSaveChanges.Value = False Then: ctrl.Enabled = True
    Next ctrl
    
    'change tglSaveChanges color
    If Me.tglSaveChanges.Value = True Then: tglSaveChanges.BackColor = RGB(146, 208, 80)
    If Me.tglSaveChanges.Value = False Then: tglSaveChanges.BackColor = RGB(255, 0, 0)
    
    'Data Validation
    If Me.tglSaveChanges.Value = True Then
        Me.CheckHeadersData
        Me.CheckOperationsData
        Me.CheckNotesData
    End If

End Sub

Public Sub CheckHeadersData() '<==Alter (Entire Sub, Unique)
'Check info in Headers frame for errors, before save.

    'Tracking Number editing prevention
    If Me.frameHeaders.Controls("txtH" & Range(Me.UCShown.ColumnAddress).Row).Value <> Me.UCShown.TrackingNumber Then
        'Error Message
        MsgBox "The original Tracking Number has been changed. Please set the Tracking Number back to: " & Me.UCShown.TrackingNumber & " and lock the Headers Frame before saving.", , "Data Validation Error!"
        'untoggle save changes button
        Me.tglSaveChanges.Value = False
    End If

End Sub

Public Sub CheckOperationsData() '<==Alter (Entire Sub, Unique)
'Check info in Operations frame for errors, before save.

    Dim strAllowedChars As String 'String containing all allowed characters for operations date cell text boxes.
    Dim ctrl As Control 'iterator
    Dim i As Integer 'iterator
    
    'set allowed characters
    strAllowedChars = "0123456789/"
    
    'check textbox values for invalid data
    For Each ctrl In Me.frameOperations.Controls
        'only textboxes need validation
        If TypeName(ctrl) = "TextBox" Then
            'values validation
            For i = 1 To Len(ctrl.Value)
                If InStr(1, strAllowedChars, Mid(ctrl.Value, i, 1)) = 0 Then
                    'Error Message
                    MsgBox "The textbox in userform row " & Me.frameOperations.Controls("rOps" & Int(Replace(ctrl.Name, "txtOps", ""))).Caption & " has an invalid character: '" & Mid(ctrl.Value, i, 1) & "'", , "Data Validation Error!"
                    'untoggle save changes button
                    Me.tglSaveChanges.Value = False
                    'exit for loop
                    Exit For
                End If
            Next i
        End If
    Next ctrl

End Sub

Public Sub CheckNotesData() '<==Alter (Entire Sub, Unique)
'Check info in Notes frame for errors, before save.

    Dim strAllowedChars As String 'String containing all allowed characters for operations date cell text boxes.
    Dim ctrl As Control 'temporary control object for data validation
    Dim i As Integer 'iterator
    
    'set allowed characters
    strAllowedChars = "0123456789/"
    
    'Rework / RTO Date TextBox
        'set ctrl variable
        Set ctrl = Me.frameNotes.Controls("txtN" & 2)
        'values validation
        For i = 1 To Len(ctrl.Value)
            If InStr(1, strAllowedChars, Mid(ctrl.Value, i, 1)) = 0 Then
                'Error Message
                MsgBox "The textbox in userform row " & Me.frameNotes.Controls("rN" & Int(Replace(ctrl.Name, "txtN", ""))).Caption & " has an invalid character: '" & Mid(ctrl.Value, i, 1) & "'", , "Data Validation Error!"
                'untoggle save changes button
                Me.tglSaveChanges.Value = False
                'exit for loop
                Exit For
            End If
        Next i
    
    'Last Date Seen
        'set ctrl variable
        Set ctrl = Me.frameNotes.Controls("txtN" & 4)
        'values validation
        For i = 1 To Len(ctrl.Value)
            If InStr(1, strAllowedChars, Mid(ctrl.Value, i, 1)) = 0 Then
                'Error Message
                MsgBox "The textbox in userform row " & Me.frameNotes.Controls("rN" & Int(Replace(ctrl.Name, "txtN", ""))).Caption & " has an invalid character: '" & Mid(ctrl.Value, i, 1) & "'", , "Data Validation Error!"
                'untoggle save changes button
                Me.tglSaveChanges.Value = False
                'exit for loop
                Exit For
            End If
        Next i

End Sub

Public Sub ClearUCData()

    Dim i As Integer 'iterator
    
    'reset toggle buttons
    Me.tglSaveChanges = False
    Me.tglH_EditLock = True
    
    'Clear Headers Frame
    For i = 1 To Me.UCShown.Headers.GroupSize
        'clear row titles
        Me.frameHeaders.Controls("lblH" & i).Caption = ""
        'clear values
        Me.frameHeaders.Controls("txtH" & i).Value = ""
        'clear colors
        Me.frameHeaders.Controls("txtH" & i).BackColor = RGB(255, 255, 255)
    Next i
    
    'Clear Operations Frame
    For i = 1 To Me.UCShown.NumberOfOps
        'clear row titles
        Me.frameOperations.Controls("lblOps" & i).Caption = ""
            'white fill with black lettering to reset labels
            Me.frameOperations.Controls("lblOps" & i).BackColor = RGB(255, 255, 255)
            Me.frameOperations.Controls("lblOps" & i).ForeColor = RGB(0, 0, 0)
        'clear values
        Me.frameOperations.Controls("txtOps" & i).Value = ""
        'clear colors
        Me.frameOperations.Controls("txtOps" & i).BackColor = RGB(255, 255, 255)
    Next i
    
    'Clear Notes Frame
    For i = 1 To Me.UCShown.Notes.GroupSize
        'clear row titles
        Me.frameNotes.Controls("lblN" & i).Caption = ""
        'clear values
        Me.frameNotes.Controls("txtN" & i).Value = ""
        'clear colors
        Me.frameNotes.Controls("txtN" & i).BackColor = RGB(255, 255, 255)
    Next i
    
    'clear Me.UCShown
    Set Me.UCShown = Nothing

End Sub

Public Sub ReadUCData(ByVal targetUC_TNumAbbr As String)
'Read data from ArrWIP to populate userform and reset Data Validation Booleans.

    Dim i As Integer 'iterator
    Dim ctrl As Control 'iterator
    
    
    'iterate arrwip for target UC
    For i = 1 To UBound(ArrWIP)
        If ArrWIP(i).TNumAbbr = targetUC_TNumAbbr Then
            Set Me.UCShown = ArrWIP(i)
            Exit For
        End If
    Next i
    
    
    'Fill Headers Frame
    For i = 1 To Me.UCShown.Headers.GroupSize
        'fill row titles
        Me.frameHeaders.Controls("lblH" & i).Caption = Me.UCShown.Headers.TitlesList(i)
        'fill values
        Me.frameHeaders.Controls("txtH" & i).Value = Me.UCShown.Headers.ValuesList(i)
        'fill colors
        Me.frameHeaders.Controls("txtH" & i).BackColor = Me.UCShown.Headers.ColorsList(i)
        'Lock all Headers Frame text boxes initially
        Me.frameHeaders.Controls("txtH" & i).Locked = True
    Next i
    
    
    'Fill Operations Frame
    For i = 1 To Me.UCShown.NumberOfOps
        'fill row titles
        Me.frameOperations.Controls("lblOps" & i).Caption = Me.UCShown.OperationsList(i).Title
            'black fill with white lettering and textbox & color button disabled if row is hidden
            If Not Me.UCShown.OperationsList(i).Enabled Then
                Me.frameOperations.Controls("lblOps" & i).BackColor = RGB(0, 0, 0)
                Me.frameOperations.Controls("lblOps" & i).ForeColor = RGB(255, 255, 255)
                Me.frameOperations.Controls("lblOps" & i).Enabled = False
                Me.frameOperations.Controls("txtOps" & i).Enabled = False
                Me.frameOperations.Controls("btnOps" & i).Enabled = False
            End If
        'fill values (ignore 0-Date values)
        If Me.UCShown.OperationsList(i).UCDate <> CDate(0) Then
            'fill textbox with date minus the year
            Me.frameOperations.Controls("txtOps" & i).Value = Left(CStr(Me.UCShown.OperationsList(i).UCDate), (Len(CStr(Me.UCShown.OperationsList(i).UCDate)) - 5))
        End If
        'fill colors
        Me.frameOperations.Controls("txtOps" & i).BackColor = Me.UCShown.OperationsList(i).UCColor
    Next i
    
    'Fill Notes Frame
    For i = 1 To Me.UCShown.Notes.GroupSize
        'fill row titles
        Me.frameNotes.Controls("lblN" & i).Caption = Me.UCShown.Notes.TitlesList(i)
        'fill values
        Me.frameNotes.Controls("txtN" & i).Value = Me.UCShown.Notes.ValuesList(i)
        'fill colors
        Me.frameNotes.Controls("txtN" & i).BackColor = Me.UCShown.Notes.ColorsList(i)
    Next i
    
    'initialize data validation booleans
    Me.bValidatedHeaders = False
    Me.bValidatedOperations = False
    Me.bValidatedNotes = False

End Sub

Public Sub InitializeControls()

    Dim ucdHeight As Integer 'Height of the UserForm (the top Caption bar is exactly 21px tall)
    Dim headersHeight As Integer 'Height of Headers Frame
    Dim headersTop As Integer 'Top of Headers Frame
    Dim opsHeight As Integer 'Height of Operations Frame
    Dim opsTop As Integer 'Top of Operations Frame
    Dim notesHeight As Integer 'Height of Notes Frame
    Dim notesTop As Integer 'Top of Notes Frame
    Dim fSpace As Integer 'space between frames
    Dim hControlTop As Integer 'current frame control top location
    Dim opsControlTop As Integer 'current frame control top location
    Dim nControlTop As Integer 'current frame control top location
    Dim scrCtrlRowNum As Integer 'current number of control rows to assign to scrollbar max value
    Dim i As Integer 'iterator
    
    'initialize frame sizes
    fSpace = 6
    headersHeight = 30
    opsHeight = 30
    notesHeight = 30
    headersTop = 48
    opsTop = headersTop + headersHeight + fSpace
    notesTop = opsTop + opsHeight + fSpace
    ucdHeight = 21 + notesTop + notesHeight + fSpace '21 is for top caption bar
    
    'initialize control top counters
    hControlTop = 6
    opsControlTop = 6
    nControlTop = 6
    
    'initialize control row number
    scrCtrlRowNum = 0
    
    'resize frames
    Me.ResizeUserForm ucdHeight
    Me.ResizeFrames "frameHeaders", headersHeight, headersTop
    Me.ResizeFrames "frameOperations", opsHeight, opsTop
    Me.ResizeFrames "frameNotes", notesHeight, notesTop
    
    'Fill Headers Frame (every row is a row "label" and a text box (h = 18px ; w = 72px))
        'place each text box at left 30
        'place each successive control at top +24 of previous
        For i = 1 To ArrWIP(1).Headers.GroupSize
            'increment control row counter
            scrCtrlRowNum = scrCtrlRowNum + 1
            'add row label, find placement, and set properties
            Me.frameHeaders.Controls.Add "Forms.Label.1", "rH" & i, True
             With Me.frameHeaders.Controls("rH" & scrCtrlRowNum)
                .Top = hControlTop + 3
                .Left = 6
                .Caption = Str(scrCtrlRowNum)
                .TextAlign = fmTextAlignCenter
                .Width = 12
            End With
            'add row title label and find placement
            Me.frameHeaders.Controls.Add "Forms.Label.1", "lblH" & i, True
            With Me.frameHeaders.Controls("lblH" & i)
                .Top = hControlTop + 3
                .Left = 24
                .TextAlign = fmTextAlignLeft
                .BackColor = RGB(255, 255, 255)
                .Width = 80
                .Height = 12
            End With
            'add textbox and find placement
            Me.frameHeaders.Controls.Add "Forms.TextBox.1", "txtH" & i, True
            With Me.frameHeaders.Controls("txtH" & i)
                .Top = hControlTop
                .Left = 110
                .Width = 104
            End With
            hControlTop = hControlTop + 18 + 6
            'increment height and top variables
            headersHeight = 6 + hControlTop
            headersTop = 48
            opsTop = headersTop + headersHeight + fSpace
            notesTop = opsTop + opsHeight + fSpace
            ucdHeight = 21 + notesTop + notesHeight + fSpace '21 is for top caption bar
            'resize frames and userform
            Me.ResizeUserForm ucdHeight
            Me.ResizeFrames "frameHeaders", headersHeight, headersTop
            Me.ResizeFrames "frameOperations", opsHeight, opsTop
            Me.ResizeFrames "frameNotes", notesHeight, notesTop
        Next i
    
    'Fill Operations Frame (every row is a row "label", a text box, and a combo box (h = 18px ; w = 72px))
        'place each text box at left 30 and combobox at left 120
        'place each successive control at top +24 of previous
        For i = 1 To ArrWIP(1).NumberOfOps
            'increment control row counter
            scrCtrlRowNum = scrCtrlRowNum + 1
            'add row label, find placement, and set caption
            Me.frameOperations.Controls.Add "Forms.Label.1", "rOps" & i, True
            With Me.frameOperations.Controls("rOps" & i)
                .Top = opsControlTop + 3
                .Left = 6
                .Caption = Str(scrCtrlRowNum)
                .TextAlign = fmTextAlignCenter
                .Width = 12
            End With
            'add row op label and find placement
            Me.frameOperations.Controls.Add "Forms.Label.1", "lblOps" & i, True
            With Me.frameOperations.Controls("lblOps" & i)
                .Top = opsControlTop + 3
                .Left = 24
                .TextAlign = fmTextAlignLeft
                .BackColor = RGB(255, 255, 255)
                .Width = 80
                .Height = 12
            End With
            'add textbox and find placement
            Me.frameOperations.Controls.Add "Forms.TextBox.1", "txtOps" & i, True
            With Me.frameOperations.Controls("txtOps" & i)
                .Top = opsControlTop
                .Left = 110
                .TextAlign = fmTextAlignRight
                .Width = 36
            End With
            'add button and find placement
            Me.frameOperations.Controls.Add "Forms.CommandButton.1", "btnOps" & i, True
            With Me.frameOperations.Controls("btnOps" & i)
                .Top = opsControlTop
                .Left = 152
                .Height = 18
                .Width = 60
                .Caption = "Color Options"
            End With
            opsControlTop = opsControlTop + 18 + 6
            'increment height and top variables
            opsHeight = 6 + opsControlTop
            opsTop = headersTop + headersHeight + fSpace
            notesTop = opsTop + opsHeight + fSpace
            ucdHeight = 21 + notesTop + notesHeight + fSpace '21 is for top caption bar
            'resize frames and userform
            Me.ResizeUserForm ucdHeight
            Me.ResizeFrames "frameOperations", opsHeight, opsTop
            Me.ResizeFrames "frameNotes", notesHeight, notesTop
        Next i
    
    'Fill Notes Frame (every row is a row "label" and a text box (h = 18px ; w = 72px))
        'place each text box at left 30
        'place each successive control at top +24 of previous
        For i = 1 To ArrWIP(1).Notes.GroupSize
            'increment control row counter
            scrCtrlRowNum = scrCtrlRowNum + 1
            'add row label, find placement, and set caption
            Me.frameNotes.Controls.Add "Forms.Label.1", "rN" & i, True
            With Me.frameNotes.Controls("rN" & i)
                .Top = nControlTop + 3
                .Left = 6
                .Caption = Str(scrCtrlRowNum)
                .TextAlign = fmTextAlignCenter
                .Width = 12
            End With
            'add row title label and find placement
            Me.frameNotes.Controls.Add "Forms.Label.1", "lblN" & i, True
            With Me.frameNotes.Controls("lblN" & i)
                .Top = nControlTop + 3
                .Left = 24
                .TextAlign = fmTextAlignLeft
                .BackColor = RGB(255, 255, 255)
                .Width = 80
                .Height = 12
            End With
            'add textbox and find placement
            Me.frameNotes.Controls.Add "Forms.TextBox.1", "txtN" & i, True
            With Me.frameNotes.Controls("txtN" & i)
                .Top = nControlTop
                .Left = 110
                .Width = 104
            End With
            nControlTop = nControlTop + 18 + 6
            'increment height and top variables
            notesHeight = 6 + nControlTop 'top won't be changed after both frames above are already set
            ucdHeight = 21 + notesTop + notesHeight + fSpace '21 is for top caption bar
            'resize frames and userform
            Me.ResizeUserForm ucdHeight
            Me.ResizeFrames "frameNotes", notesHeight, notesTop
        Next i
    
    'create scrollbars (ucdHeight - 21) gives perfect distance for scrolling
    Me.SetScrollBar (ucdHeight - 21) 'subtract the 21px top caption bar

End Sub

Public Sub SetScrollBar(ByVal scrlHeight As Single)

    With Me
        'This will create a vertical scrollbar
        .ScrollBars = fmScrollBarsVertical
        
        'Change the values of 2 as Per your requirements
        .ScrollHeight = scrlHeight
    End With

End Sub

Public Sub ResizeFrames(ByVal frameTitle As String, ByVal frameHeight As Integer, ByVal frameTop As Integer)

    Me.Controls(frameTitle).Height = frameHeight
    Me.Controls(frameTitle).Top = frameTop

End Sub

Public Sub ResizeUserForm(ByVal formHeight As Integer)

    Dim yScreen As Long 'holds screen height
    
    'assign screen height variable from function
    yScreen = (GetSystemMetrics(SM_CYSCREEN) * 0.7) '70% seems to be the perfect size for any screen
    
    'prevent userform from receiving height greater than screen height
    If formHeight >= yScreen Then
        Me.Height = yScreen
    Else
        Me.Height = formHeight
    End If

End Sub

Public Sub StartScrollHandler() 'Fires every time the userform is shown

'!!!!!====Tie MouseWheel to ONLY UCDisplay until removed with unhook====!!!!!'
    HookFormScroll Me
'!!!!!====Tie MouseWheel to ONLY UCDisplay until removed with unhook====!!!!!'

End Sub

Public Sub StopScrollHandler()

'!!!!!====Remove MouseWheel from this form====!!!!!'
    UnhookFormScroll
'!!!!!====Remove MouseWheel from this form====!!!!!'

End Sub

Private Sub UserForm_Initialize()

    Dim ucdEvent As TC_EH_UCDisplay 'UC Display Event Handler
    Dim i As Integer 'iterator
    
    'initialize controls on form
    Me.InitializeControls
    
    'initialize event handlers
        'Headers
        For i = 1 To ArrWIP(1).Headers.GroupSize
            Set ucdEvent = New TC_EH_UCDisplay
            Set ucdEvent.frmRef_UCDisplay = Me
            Set ucdEvent.txtRef_UCDisplay = Me.frameHeaders.Controls("txtH" & i)
            ucdEvent.rowIndex_UCDisplay = i
            'Add to Event Handler Collection
            cEH_UCDisplay.Add ucdEvent
        Next i
        'Operations
        For i = 1 To ArrWIP(1).NumberOfOps
            Set ucdEvent = New TC_EH_UCDisplay
            Set ucdEvent.frmRef_UCDisplay = Me
            Set ucdEvent.btnRef_UCDisplay = Me.Controls("btnOps" & i)
            Set ucdEvent.txtRef_UCDisplay = Me.frameOperations.Controls("txtOps" & i)
            ucdEvent.btnIndex_UCDisplay = i
            ucdEvent.rowIndex_UCDisplay = i + ArrWIP(1).Headers.GroupSize
            'Add to Event Handler Collection
            cEH_UCDisplay.Add ucdEvent
        Next i
        'Notes
        For i = 1 To ArrWIP(1).Notes.GroupSize
            Set ucdEvent = New TC_EH_UCDisplay
            Set ucdEvent.frmRef_UCDisplay = Me
            Set ucdEvent.txtRef_UCDisplay = Me.frameNotes.Controls("txtN" & i)
            ucdEvent.rowIndex_UCDisplay = i + ArrWIP(1).Headers.GroupSize + ArrWIP(1).NumberOfOps
            'Add to Event Handler Collection
            cEH_UCDisplay.Add ucdEvent
        Next i

End Sub

Private Sub UserForm_Activate() 'Fires every time the userform is shown

    'call scroll handler
    Me.StartScrollHandler

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent userform close on red x click.

    If CloseMode = 0 Then: Cancel = True

End Sub
