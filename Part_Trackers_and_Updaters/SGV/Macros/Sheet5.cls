VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ToggleButton22_Click()
If ToggleButton22.Value = True Then
 'This area contains the things you want to happen
'when the toggle button is not depressed
Rows(41).EntireRow.Hidden = True
Rows(42).EntireRow.Hidden = True
Rows(43).EntireRow.Hidden = True
Rows(44).EntireRow.Hidden = True
Rows(62).EntireRow.Hidden = True
Rows(63).EntireRow.Hidden = True
Rows(64).EntireRow.Hidden = True
Rows(65).EntireRow.Hidden = True
Rows(83).EntireRow.Hidden = True
Rows(84).EntireRow.Hidden = True
Rows(85).EntireRow.Hidden = True
Rows(86).EntireRow.Hidden = True
Rows(104).EntireRow.Hidden = True
Rows(105).EntireRow.Hidden = True
Rows(106).EntireRow.Hidden = True
Rows(107).EntireRow.Hidden = True
Rows(125).EntireRow.Hidden = True
Rows(126).EntireRow.Hidden = True
Rows(127).EntireRow.Hidden = True
Rows(128).EntireRow.Hidden = True
Rows(146).EntireRow.Hidden = True
Rows(147).EntireRow.Hidden = True
Rows(148).EntireRow.Hidden = True
Rows(149).EntireRow.Hidden = True
Rows(167).EntireRow.Hidden = True
Rows(168).EntireRow.Hidden = True
Rows(169).EntireRow.Hidden = True
Rows(170).EntireRow.Hidden = True

Else
'This area contains the things you want to happen
'when the toggle button is depressed
Rows(41).EntireRow.Hidden = False
Rows(42).EntireRow.Hidden = False
Rows(43).EntireRow.Hidden = False
Rows(44).EntireRow.Hidden = False
Rows(62).EntireRow.Hidden = False
Rows(63).EntireRow.Hidden = False
Rows(64).EntireRow.Hidden = False
Rows(65).EntireRow.Hidden = False
Rows(83).EntireRow.Hidden = False
Rows(84).EntireRow.Hidden = False
Rows(85).EntireRow.Hidden = False
Rows(86).EntireRow.Hidden = False
Rows(104).EntireRow.Hidden = False
Rows(105).EntireRow.Hidden = False
Rows(106).EntireRow.Hidden = False
Rows(107).EntireRow.Hidden = False
Rows(125).EntireRow.Hidden = False
Rows(126).EntireRow.Hidden = False
Rows(127).EntireRow.Hidden = False
Rows(128).EntireRow.Hidden = False
Rows(146).EntireRow.Hidden = False
Rows(147).EntireRow.Hidden = False
Rows(148).EntireRow.Hidden = False
Rows(149).EntireRow.Hidden = False
Rows(167).EntireRow.Hidden = False
Rows(168).EntireRow.Hidden = False
Rows(169).EntireRow.Hidden = False
Rows(170).EntireRow.Hidden = False

 End If
End Sub
