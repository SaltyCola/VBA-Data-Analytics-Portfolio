VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TC_UCSearch 
   Caption         =   "Unit Column Search:"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "TC_UCSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TC_UCSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnMainMenu_Enter()
'Loads main menu uf and closes UC Search.

    Dim i As Integer 'iterator
    
    'reset all wSorted booleans for waterfalling later (used in this userform as a way to prevent multiple adds to the listbox)
    For i = 1 To UBound(ArrWIP)
        ArrWIP(i).WSorted = False
    Next i
    
    'reset ArrListBox
    Erase ArrListBox()
    
    'load Main Menu
    Me.LoadMainMenu

End Sub

Private Sub btnSearch_Click()
'Searches ArrWIP for correct UC. Single search uses textbox value while multi uses first
'listbox entry and will apply any changes made to all UCs.

    Dim i As Integer 'iterator
    Dim strSearch As String 'search string variable
    Dim bFormat As Boolean 'True: strSearch is the correct format ; False: not correct format
    
    'reset all wSorted booleans for waterfalling later (used in this userform as a way to prevent multiple adds to the listbox)
    For i = 1 To UBound(ArrWIP)
        ArrWIP(i).WSorted = False
    Next i
    
    'check for correct format ("J0001")
    strSearch = UCase(Me.txtSearch.Value)
    bFormat = SearchStringFormat(strSearch)
    
    'incorrect formatting (only if single search option selected)
    If Not bFormat And Me.optSingle.Value Then
        MsgBox "The Search String must be one letter, followed by 4 numbers." & vbNewLine & "(i.e. j0001 or J0001)", , "Incorrect Search Format" '<==Alter
        'send focus to textbox
        Me.txtSearch.SetFocus
    
    'option single is chosen
    ElseIf Me.optSingle.Value Then
        'check for UC in ArrWIP
        For i = 1 To UBound(ArrWIP)
            'UC found
            If ArrWIP(i).TNumAbbr = strSearch Then
                'load UCDisplay
                Me.LoadUCDisplay ArrWIP(i)
                Exit For
            'UC not found
            ElseIf i = UBound(ArrWIP) And ArrWIP(i).TNumAbbr <> strSearch Then
                MsgBox "This Unit Column was not found in WIP..."
                'send focus to textbox
                Me.txtSearch.SetFocus
            End If
        Next i
    
    'option multiple is chosen and listbox isn't empty
    ElseIf Me.optMultiple.Value And Me.lstMult.ListCount <> 0 Then
        'Warning Message
        MsgBox "ANY changes made on the next window will be applied to all Unit Columns listed...", , "WARNING"
        'change multi-update boolean
        ufUCDisplay.bMultiUpdate = True
        'load UCDisplay for first entry (other entries will be updated with only the changes made to the first)
        Me.LoadUCDisplay ArrListBox(1)
    
    'option multiple is chosen and listbox is empty
    ElseIf Me.optMultiple.Value And Me.lstMult.ListCount = 0 Then
        MsgBox "Listbox has no entries. Please add entries by pressing enter while in the textbox."
        'send focus to textbox
        Me.txtSearch.SetFocus
    
    End If

End Sub

Private Sub optSingle_Change()
'Tests for an option selection change between single search or multi-search.

    'Disable listbox while single option is clicked
    If Me.optSingle.Value = True Then
        Me.lstMult.ForeColor = RGB(200, 200, 200)
    ElseIf Me.optSingle.Value = False Then
        Me.lstMult.ForeColor = RGB(0, 0, 0)
    End If
    
    'send focus to textbox
    Me.txtSearch.SetFocus

End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As msforms.ReturnInteger, ByVal Shift As Integer)
'Directs Userform what to do when the enter key is pressed while focus is on the textbox.

    Dim i As Integer 'iterator
    Dim strSearch As String 'search string variable
    Dim bFormat As Boolean 'True: strSearch is the correct format ; False: not correct format
    
    
    'enter is pressed while in the text box
    If KeyCode = vbKeyReturn Then
    
        'cancel enter key to keep focus
        KeyCode = 0
    
        'assign search string
        strSearch = UCase(Me.txtSearch.Value)
        'check for correct format ("J0001")
        bFormat = SearchStringFormat(strSearch)
        
        
        'incorrect format
        If Not bFormat Then
            'incorrect formatting message
            MsgBox "The Search String must be one letter, followed by 4 numbers." & vbNewLine & "(i.e. j0001 or J0001)", , "Incorrect Search Format" '<==Alter
        
        'single UC being searched
        ElseIf Me.optSingle Then
            Me.btnSearch.SetFocus
        
        
        'multiple UC being searched
        ElseIf Me.optMultiple Then
            'add UC to list box if it exists and is not already in the listbox
            For i = 1 To UBound(ArrWIP)
                If ArrWIP(i).TNumAbbr = strSearch And Not ArrWIP(i).WSorted Then 'UC found in ArrWIP and hasn't been added already (wSorted)
                    'add entry to listbox
                    Me.lstMult.AddItem ArrWIP(i).TrackingNumber
                    'activate waterfall sorted boolean to prevent multiple adds
                    ArrWIP(i).WSorted = True
                     'resize array
                    ReDim Preserve ArrListBox(1 To Me.lstMult.ListCount) As TC_UnitColumn
                    'add UC to array
                    Set ArrListBox(Me.lstMult.ListCount) = ArrWIP(i)
                    'clear textbox
                    Me.txtSearch.Value = ""
                    Exit For
                ElseIf ArrWIP(i).TNumAbbr = strSearch And ArrWIP(i).WSorted Then 'UC already added to listbox
                    MsgBox "This Unit Column was already added to the list."
                ElseIf i = UBound(ArrWIP) Then 'only fires if UC was not found at the last possible UC searched
                    MsgBox "This Unit Column was not found in WIP..."
                End If
            Next i
        End If
    
    
    End If


End Sub

Private Function SearchStringFormat(ByVal strSearch As String) As Boolean
'Verifies whether the text entered into the textbox follows the correct format.

    Dim rtnBool As Boolean 'return boolean
    Dim i As Integer 'iterator
    
    'initialize rtnBool
    rtnBool = True
    
    'first char is a letter
    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(strSearch, 1)) = 0 Then: rtnBool = False '<==Alter
    
    'last four char are numbers
    For i = 2 To 5
        If InStr(1, "0123456789", Mid(strSearch, i, 1)) = 0 Then: rtnBool = False '<==Alter
    Next i
    
    'must have 5 char total
    If Not Len(strSearch) = 5 Then: rtnBool = False '<==Alter
    
    'return value
    SearchStringFormat = rtnBool

End Function

Public Sub LoadUCDisplay(ByVal displayUnitColumn As TC_UnitColumn)
'Prepares ufUCDisplay to be shown with the correct information, as well as clears out
'ufUCSearch for future use.

    'select column in Tracker
    Application.Goto SheetWIP.Range(displayUnitColumn.ColumnAddress), Scroll:=True
    
    'hide and clear search userform
    Me.Hide
    Me.txtSearch.Value = ""
    Me.optSingle.Value = True
    Me.optMultiple.Value = False
    Me.lstMult.Clear
    
    'show UCDisplay (without tracker movement)
    ufUCDisplay.ReadUCData (displayUnitColumn.TNumAbbr)
    ufUCDisplay.Show

End Sub

Public Sub LoadMainMenu()
'Prepares Main Menu to be shown, as well as clears out ufUCSearch for future use.

    'hide and clear search userform
    Me.Hide
    Me.txtSearch.Value = ""
    Me.optSingle.Value = True
    Me.optMultiple.Value = False
    Me.lstMult.Clear
    
    'show Main Menu (without tracker movement)
    ufMainMenu.Show

End Sub

Private Sub UserForm_Activate()
'Focus on textbox every time the userform is shown.

    Me.txtSearch.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent userform close on red x click.

    If CloseMode = 0 Then: Cancel = True

End Sub
