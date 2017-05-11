VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_DeleteError 
   Caption         =   "Cell Error Found"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   OleObjectBlob   =   "SN_DeleteError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_DeleteError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub ClearCellButton_Click()

    Worksheets("NEO 5322121").Range(rngErrorCell) = ""
    boolFlagCellClear = True
    Me.Hide

End Sub

Private Sub FlagCellButton_Click()

    boolFlagCell = True
    Me.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        boolCanceled = True
        Me.Hide
        frmSNInfoPage.Hide
    End If

End Sub
