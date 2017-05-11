VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ES_EngineSetUpdater 
   Caption         =   "Update Engine Set Progress"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   OleObjectBlob   =   "ES_EngineSetUpdater.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ES_EngineSetUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub TextBoxESU_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        
        Dim txtEngSet As String
        Dim i As Integer
        Dim rngEngSet As Range
        Dim res As Range
        Dim rngESCol As Range
        Dim cllCol As Range
        Dim boolFirstColoredCell As Boolean
        
        'initialize booleans
        boolAllowToggles = False
        boolFirstColoredCell = True
        
        'define ranges
        Set rngEngSet = Worksheets("NEO 5322121 Aggressive LTs").Range("1:1")
        
        'define search text
        txtEngSet = TextBoxESU.Value
        
            'Engine Set number format errors
        'Length error
        If Not Len(txtEngSet) = 6 Then
            'length error message
            MsgBox "Please enter a 6-digit number.", , "Length Error"
            TextBoxESU.SetFocus
        'Alphanumeric error
        ElseIf (Len(txtEngSet) = 6) And (Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 1, 1)) = 0) Or Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 2, 1)) = 0) Or Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 3, 1)) = 0) Or Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 4, 1)) = 0) Or Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 5, 1)) = 0) Or Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(txtEngSet, 6, 1)) = 0)) Then
            'alphanumeric error message
            MsgBox "The Engine Set number must only contain numbers.", , "Format Error"
            TextBoxESU.SetFocus
        
        'correct format
        Else
    
            'iterate top row for columns
            For Each res In rngEngSet.SpecialCells(xlCellTypeVisible)
            
                'Eng Set not found when black line reached
                If res.Interior.Color = RGB(0, 0, 0) Then
                    'eng set not found message
                    MsgBox "That Engine Set number can not be found...", , "Search Error"
                    TextBoxESU.SetFocus
                    Exit Sub
                End If
            
                'Eng Set number found
                If res.Value = txtEngSet Then
                
                    'define engine set column
                    esCol = res.Column
                    
                    'search column for colors
                    Set rngESCol = Worksheets("NEO 5322121 Aggressive LTs").Range(Cells(7, res.Column), Cells(33, res.Column))
                    
                    'iterate row
                    For Each cllCol In rngESCol
                        
                        'white cell found
                        If cllCol.Interior.Color = clrWhite Then
                            Me.Controls("ToggleR" & cllCol.row).Value = False
                            Me.Controls("ToggleR" & cllCol.row).BackColor = cllCol.Interior.Color
                            Me.Controls("ToggleR" & cllCol.row).Enabled = True
                            TextBoxESU.SetFocus
                        'first non white found
                        ElseIf (boolFirstColoredCell) And ((cllCol.Interior.Color = clrGreen) Or (cllCol.Interior.Color = clrYellow) Or (cllCol.Interior.Color = clrRed)) Then
                            Me.Controls("ToggleR" & cllCol.row).Value = True
                            Me.Controls("ToggleR" & cllCol.row).BackColor = cllCol.Interior.Color
                            Me.Controls("ToggleR" & cllCol.row).Enabled = True
                            boolFirstColoredCell = False
                            TextBoxESU.SetFocus
                        'all other non whites found
                        ElseIf Not (boolFirstColoredCell) And ((cllCol.Interior.Color = clrGreen) Or (cllCol.Interior.Color = clrYellow) Or (cllCol.Interior.Color = clrRed)) Then
                            Me.Controls("ToggleR" & cllCol.row).Value = True
                            Me.Controls("ToggleR" & cllCol.row).BackColor = cllCol.Interior.Color
                            Me.Controls("ToggleR" & cllCol.row).Enabled = False
                            TextBoxESU.SetFocus
                        End If
                    
                    Next cllCol
                    
                    'reset boolean
                    boolAllowToggles = True
                    
                    'exit sub to stop searching
                    Exit Sub
                
                End If
            Next res
        End If
    KeyCode = 0
    End If

End Sub

Private Sub ToggleR7_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR7.Value
        esRow = 7
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR8_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR8.Value
        esRow = 8
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR9_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR9.Value
        esRow = 9
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR10_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR10.Value
        esRow = 10
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR11_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR11.Value
        esRow = 11
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR12_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR12.Value
        esRow = 12
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR13_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR13.Value
        esRow = 13
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR14_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR14.Value
        esRow = 14
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR15_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR15.Value
        esRow = 15
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR16_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR16.Value
        esRow = 16
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR17_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR17.Value
        esRow = 17
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR18_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR18.Value
        esRow = 18
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR19_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR19.Value
        esRow = 19
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR20_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR20.Value
        esRow = 20
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR21_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR21.Value
        esRow = 21
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR22_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR22.Value
        esRow = 22
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR23_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR23.Value
        esRow = 23
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR24_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR24.Value
        esRow = 24
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR25_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR25.Value
        esRow = 25
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR26_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR26.Value
        esRow = 26
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR27_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR27.Value
        esRow = 27
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR28_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR28.Value
        esRow = 28
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR29_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR29.Value
        esRow = 29
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR30_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR30.Value
        esRow = 30
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR31_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR31.Value
        esRow = 31
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR32_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR32.Value
        esRow = 32
        Call ToggleButtonHandler
    End If

End Sub

Private Sub ToggleR33_Click()

    'if manually toggling buttons
    If boolAllowToggles Then
    
        'define variables
        boolToggled = Me.ToggleR33.Value
        esRow = 33
        Call ToggleButtonHandler
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        'Cancel = True
    End If

End Sub


