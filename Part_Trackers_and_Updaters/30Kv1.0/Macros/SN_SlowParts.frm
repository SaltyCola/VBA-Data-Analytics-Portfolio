VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_SlowParts 
   Caption         =   "Slow Part Analyzer"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3480
   OleObjectBlob   =   "SN_SlowParts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_SlowParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub ListBox1_Click()

    '==============================Search Tracker==============================='
    Dim trgtStr As String
    Dim snRng As Range
    Dim sn As Range
    Dim sncolRng As Range
    Dim sncolCell As Range
    
    'Define snRng
    Set snRng = Worksheets("NEO 5322121").Range("6:6")
    
    'Define variables
    trgtStr = ListBox1.Text
    
    'search sn row in tracker
    For Each sn In snRng
        
        'serial number found
        If sn.Value = trgtStr Then
            
            'search for first non blank colored cell
            Set sncolRng = Worksheets("NEO 5322121").Range(Cells(7, sn.Column), Cells(43, sn.Column))
            For Each sncolCell In sncolRng
                If Not sncolCell.Interior.Color = clrBlank Then
                    'go to location
                    Application.Goto sncolCell.Offset(-5, 0), Scroll:=True
                    'enable slow toggle button
                    ToggleSlowPart.Enabled = True
                    'fill last date seen textboxes
                    TextBoxLastDateSeen.Value = Worksheets("NEO 5322121").Cells(52, sn.Column).Value
                    TextBoxDaysSince.Value = (Date - Worksheets("NEO 5322121").Cells(52, sn.Column).Value)
                    Exit For
                End If
            Next sncolCell
            
        End If
        
    Next sn
    '==========================================================================='

End Sub

Private Sub ToggleSlowPart_Click()

    'if toggled
    If ToggleSlowPart.Value = True Then
        Worksheets("NEO 5322121").Cells(6, ActiveCell.Column).Interior.Color = RGB(244, 158, 228)
    
    'if untoggled
    ElseIf ToggleSlowPart.Value = False Then
        Worksheets("NEO 5322121").Cells(6, ActiveCell.Column).Interior.Color = clrBlank
    
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        boolCanceled = True
    End If

End Sub

