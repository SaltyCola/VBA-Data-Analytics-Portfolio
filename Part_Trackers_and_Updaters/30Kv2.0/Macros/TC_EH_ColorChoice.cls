VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_EH_ColorChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TC_ColorChoice Toggle Button Event Handler Class
Public WithEvents tglRef_ColorChoice As msforms.ToggleButton 'toggle button reference
Attribute tglRef_ColorChoice.VB_VarHelpID = -1
Public frmRef_ColorChoice As UserForm 'userform reference
Public tglIndex_ColorChoice As Integer 'toggle button index reference
Public txtRef_from_UCDisplay As msforms.TextBox 'textbox reference from UCDisplay Form

Private Sub tglRef_ColorChoice_Click()
'Check for which button is toggled and untoggle the rest

    Dim tcCellColors As TC_CellColor_30K 'cell colors object to paint colors to buttons
    Dim i As Integer 'iterator
    
    'initialize cell color object
    Set tcCellColors = New TC_CellColor_30K
    
    
    'is button toggled?
    If tglRef_ColorChoice.Value = True Then
    
        'lock toggled button
        tglRef_ColorChoice.Locked = True
        
        'untoggle and unlock every other button
        For i = 0 To (tcCellColors.NumberOfColors - 1)
            If i <> tglIndex_ColorChoice Then
                frmRef_ColorChoice.Controls("ToggleButton" & i).Value = False
                frmRef_ColorChoice.Controls("ToggleButton" & i).Locked = False
            End If
        Next i
        
        'Color corresponding textbox on UCDisplay
        txtRef_from_UCDisplay.BackColor = tglRef_ColorChoice.BackColor
    
    End If
    

End Sub
