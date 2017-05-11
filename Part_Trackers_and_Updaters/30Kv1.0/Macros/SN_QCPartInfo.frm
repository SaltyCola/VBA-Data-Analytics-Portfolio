VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_QCPartInfo 
   Caption         =   "QC Part Information"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "SN_QCPartInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_QCPartInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub ToggleButtonBad_Click()

    'change txtbox and QC Status Cell color
    If ToggleButtonBad.Value = True Then
        frmSNQCPartInfo.TextBoxQCStatus.BackColor = RGB(255, 192, 0)
        frmSNQCPartInfo.TextBoxQCStatus.Value = "Bad"
        Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color = RGB(255, 192, 0)
        Worksheets("NEO 5322121").Cells(54, snSearchCol).Value = "Bad"
    ElseIf ToggleButtonBad.Value = False Then
        frmSNQCPartInfo.TextBoxQCStatus.BackColor = clrBlank
        frmSNQCPartInfo.TextBoxQCStatus.Value = ""
        Worksheets("NEO 5322121").Cells(54, snSearchCol).Interior.Color = clrBlank
        Worksheets("NEO 5322121").Cells(54, snSearchCol).Value = ""
    End If

End Sub

Private Sub ButtonConfirm_Click()

    'hide form
    Me.Hide

End Sub

Private Sub TextBoxQCStatus_Enter()

    Application.Goto Worksheets("NEO 5322121").Cells(54, snSearchCol)

End Sub

Private Sub TextBoxQCStatus_Change()

    Worksheets("NEO 5322121").Cells(54, snSearchCol).Value = Me.TextBoxQCStatus.Value

End Sub

Private Sub TextBoxRiskProfile_Enter()

    Application.Goto Worksheets("NEO 5322121").Cells(56, snSearchCol)

End Sub

Private Sub TextBoxRiskProfile_Change()

    Worksheets("NEO 5322121").Cells(56, snSearchCol).Value = Me.TextBoxRiskProfile.Value

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If

End Sub
