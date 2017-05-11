VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ToggleButton21_Click()
If ToggleButton21.Value = True Then
 'This area contains the things you want to happen
'when the toggle button is not depressed

Rows(10).EntireRow.Hidden = True
Rows(11).EntireRow.Hidden = True
Rows(12).EntireRow.Hidden = True
Rows(13).EntireRow.Hidden = True

Else
'This area contains the things you want to happen
'when the toggle button is depressed
Rows(10).EntireRow.Hidden = False
Rows(11).EntireRow.Hidden = False
Rows(12).EntireRow.Hidden = False
Rows(13).EntireRow.Hidden = False

End If
End Sub

Private Sub ToggleButton22_Click()
If ToggleButton22.Value = True Then
 'This area contains the things you want to happen
'when the toggle button is not depressed
Rows(16).EntireRow.Hidden = True
Rows(17).EntireRow.Hidden = True
Rows(18).EntireRow.Hidden = True
Rows(19).EntireRow.Hidden = True
Rows(20).EntireRow.Hidden = True

Else
'This area contains the things you want to happen
'when the toggle button is depressed
Rows(16).EntireRow.Hidden = False
Rows(17).EntireRow.Hidden = False
Rows(18).EntireRow.Hidden = False
Rows(19).EntireRow.Hidden = False
Rows(20).EntireRow.Hidden = False

 End If
End Sub

