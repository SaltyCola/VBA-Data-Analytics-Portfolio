VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MsgWaterfallETA 
   Caption         =   "Continue?"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   OleObjectBlob   =   "MsgWaterfallETA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MsgWaterfallETA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonYes_Click()

    boolWtrFllCont = True
    Me.Hide

End Sub

Private Sub ButtonNo_Click()

    boolWtrFllCont = False
    Me.Hide
    Call TrackerMainMenu

End Sub
