VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    'Unrecognized action
    On Error GoTo lineExitSub

    'Column E
    If Not Intersect(Target, Range("E:E")) Is Nothing Then
        Select Case Target
            Case Worksheets("List Source").Cells(3, 1).Value: Call GreenCell
            Case Worksheets("List Source").Cells(4, 1).Value: Call LightGreenCell
            Case Worksheets("List Source").Cells(5, 1).Value: Call WhiteCell
            Case Worksheets("List Source").Cells(6, 1).Value: Call OrangeCell
            Case Worksheets("List Source").Cells(7, 1).Value: Call YellowCell
            Case Worksheets("List Source").Cells(8, 1).Value: Call RedCell
            On Error Resume Next
        End Select
    End If
    
    'Column F
    If Not Intersect(Target, Range("F:F")) Is Nothing Then
        Select Case Target
            Case Worksheets("List Source").Cells(3, 2).Value: Call WhiteCell
            Case Worksheets("List Source").Cells(4, 2).Value: Call LightYellowCell
            Case Worksheets("List Source").Cells(5, 2).Value: Call YellowCell
            Case Worksheets("List Source").Cells(6, 2).Value: Call DarkYellowCell
            Case Worksheets("List Source").Cells(7, 2).Value: Call BlueCell
            Case Worksheets("List Source").Cells(8, 2).Value: Call RedCell
            On Error Resume Next
        End Select
    End If
    
lineExitSub:
    Exit Sub
    
End Sub





