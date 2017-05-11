VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()
'Hide 30K Detail Allocation
Range("A9:A17").EntireRow.Hidden = CheckBox1.Value
End Sub

Private Sub CheckBox2_Click()
'hide BA Detail Allocation
Range("A18:A25").EntireRow.Hidden = CheckBox2.Value
End Sub

Private Sub CheckBox3_Click()
'Hide MRJ Detail Allocation
Range("A26:A31").EntireRow.Hidden = CheckBox3.Value
End Sub

Private Sub ToggleButton2_Click()
'hide Columns which show lead time inputs
Columns("A:M").Hidden = Not Columns("A:M").Hidden
End Sub
