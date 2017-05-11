VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_WtrfllWIPorQC 
   Caption         =   "Which tab do you want to waterfall?"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "SN_WtrfllWIPorQC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_WtrfllWIPorQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonQC_Click()

    wtrfll_WIPorQC = 2
    Me.Hide

End Sub

Private Sub ButtonWIP_Click()

    wtrfll_WIPorQC = 1
    Me.Hide

End Sub
