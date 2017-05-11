VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TC_YesNoMsg 
   Caption         =   "Are you sure?"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   OleObjectBlob   =   "TC_YesNoMsg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TC_YesNoMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public bYesNoMsg As Boolean 'Boolean for continuing or not continuing code execution from userform button presses

Private Sub btnNo_Click()

    'set boolean
    Me.bYesNoMsg = False
    
    'hide form
    Me.Hide

End Sub

Private Sub btnYes_Click()

    'set boolean
    Me.bYesNoMsg = True
    
    'hide form
    Me.Hide

End Sub

Public Sub YesNoMsgInitialize(ByVal strMessage As String)

    'change label caption
    Me.Label1.Caption = strMessage
    
    'show userform
    Me.Show

End Sub

Private Sub UserForm_Activate()
'Reset yes/no boolean every time the userform is shown.

    'Always reset to False (or "No") in case of an accidental button press.
    Me.bYesNoMsg = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent userform close on red x click.

    If CloseMode = 0 Then: Cancel = True

End Sub
