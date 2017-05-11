VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_EH_UCDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TC_UCDisplay Color Button and Text Box Focus Event Handler Class
Public frmRef_UCDisplay As TC_UCDisplay 'userform reference
Public WithEvents btnRef_UCDisplay As msforms.CommandButton 'Color button reference
Attribute btnRef_UCDisplay.VB_VarHelpID = -1
Public WithEvents txtRef_UCDisplay As msforms.TextBox 'textbox referenc
Attribute txtRef_UCDisplay.VB_VarHelpID = -1
Public btnIndex_UCDisplay As Integer 'color button index reference
Public rowIndex_UCDisplay As Integer 'textbox row reference

Private Sub txtRef_UCDisplay_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Scroll SheetWIP to same row as textbox with current focus.

    Application.GoTo SheetWIP.Cells((rowIndex_UCDisplay), (frmRef_UCDisplay.UCShown.ColumnNumber))

End Sub

Private Sub btnRef_UCDisplay_Click()
'Load a ColorChoice userform overtop of button clicked

    Dim tcCellColors As TC_CellColor_24K 'cell colors object to paint colors to buttons
    Dim ccEvent As TC_EH_ColorChoice 'Toggle Button Event Handler
    Dim i As Integer 'iterator
    
    'initialize cell color object
    Set tcCellColors = New TC_CellColor_24K
    
    'initialize new color chooser
    Set ufColorChooser = Nothing
    Set ufColorChooser = New TC_ColorChoice_24K
    
    'place color chooser over colors button pressed
    ufColorChooser.StartUpPosition = 0
    ufColorChooser.Top = frmRef_UCDisplay.frameOperations.Controls("btnOps" & btnIndex_UCDisplay).Top + frmRef_UCDisplay.frameOperations.Top + 21 + frmRef_UCDisplay.Top - frmRef_UCDisplay.ScrollTop
    ufColorChooser.Left = frmRef_UCDisplay.frameOperations.Controls("btnOps" & btnIndex_UCDisplay).Left + frmRef_UCDisplay.frameOperations.Left + frmRef_UCDisplay.Left
    
    'Reset Event Handler Collection
    Set cEH_ColorChoice = New Collection
    
    'initialize event handler
    For i = 0 To (tcCellColors.NumberOfColors - 1)
        'initialize
        Set ccEvent = New TC_EH_ColorChoice
        Set ccEvent.frmRef_ColorChoice = ufColorChooser
        Set ccEvent.tglRef_ColorChoice = ufColorChooser.Controls("ToggleButton" & i)
        ccEvent.tglIndex_ColorChoice = i
        Set ccEvent.txtRef_from_UCDisplay = txtRef_UCDisplay
        'Add to Event Handler Collection
        cEH_ColorChoice.Add ccEvent
    Next i
    
    'apply color options to buttons
    For i = 0 To (tcCellColors.NumberOfColors - 1)
        ufColorChooser.Controls("ToggleButton" & i).BackColor = tcCellColors.Items(i)
    Next i
    
    'show color chooser userform (prevent actions elsewhere until closed)
    ufColorChooser.Show

End Sub
