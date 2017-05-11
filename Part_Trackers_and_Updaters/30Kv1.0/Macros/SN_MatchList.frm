VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_MatchList 
   Caption         =   "Multiple Search Results:"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "SN_MatchList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_MatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Userform_Initialize()

    Me.CommandButtonCont.SetFocus

End Sub

Private Sub CommandButtonCancel_Click()

    'hide form and clear combobox
    Dim cmbbx As Double
    For cmbbx = 1 To frmSNMatchList.ComboBox1.ListCount
        frmSNMatchList.ComboBox1.RemoveItem (0)
    Next cmbbx
    Me.Hide
    'define boolean
    boolCanceled = True
    'call snsearch
    If intSNMLType = 1 Then: Call SNSearchBox
    'call snAsBuilt
    If boolAsBuilt Then: Call SNAsBuilt
    'call snShipped
    If intSNMLType = 4 Then: Call SNShipped
    'call snQCtoWIP
    If intSNMLType = 5 Then: Call SNQCtoWIP

End Sub

Private Sub CommandButtonCont_Click()

    Dim snRng As Range
    Dim sn As Range
    Dim snCol As Long
    Dim snRow As Long
    
    'disable combobox
    frmSNMatchList.ComboBox1.Enabled = False
    
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
            
            'enable combobox
            frmSNMatchList.ComboBox1.Enabled = True
            
            'hide form and clear combobox
            Dim cmbbx As Double
            For cmbbx = 1 To frmSNMatchList.ComboBox1.ListCount
                frmSNMatchList.ComboBox1.RemoveItem (0)
            Next cmbbx
            Me.Hide
            
            'coming from snsearch
            If intSNMLType = 1 Then
                'Call info page form
                Call SNInfoPage
                Exit Sub
            'coming from snduplicate
            ElseIf intSNMLType = 2 Then
                'set first run to false
                boolSNDTGFirstRun = False
                'clear out sn duplicate textbox
                frmSNDuplicate.TextBox.Value = ""
                'call duplicate to group form
                Call SNDuplicateToGroup
                Exit Sub
            'coming from snasbuilt
            ElseIf intSNMLType = 3 Then
                'set boolean and abCol for SNML to AB
                boolSNMLtoAB = True
                abCol = snCol
                'call snAsBuilt
                Call SNAsBuilt
                Exit Sub
            'coming from snshipped
            ElseIf intSNMLType = 4 Then
                'set boolean
                boolSNMLtoShipped = True
                'call snShipped
                Call SNShipped
                Exit Sub
            'if coming from snQCtoWIP, move SN to WIP then send focus to sninfopage
            ElseIf intSNMLType = 5 Then
                frmQCtoWIP.ButtonMovetoWIP.Enabled = True
                frmQCtoWIP.ButtonMovetoWIP.SetFocus
                frmQCtoWIP.Show
                Exit Sub
            End If
            
        End If
    
    Next sn
    
    'enable combobox
    frmSNMatchList.ComboBox1.Enabled = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        
        'hide form and clear combobox
        Dim cmbbx As Double
        For cmbbx = 1 To frmSNMatchList.ComboBox1.ListCount
            frmSNMatchList.ComboBox1.RemoveItem (0)
        Next cmbbx
        Me.Hide
        
        'set boolCanceled
        boolCanceled = True
    End If

End Sub
