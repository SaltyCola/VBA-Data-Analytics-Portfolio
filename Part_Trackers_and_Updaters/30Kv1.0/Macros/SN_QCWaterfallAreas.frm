VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_QCWaterfallAreas 
   Caption         =   "Choose QC Waterfall Section(s)"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3600
   OleObjectBlob   =   "SN_QCWaterfallAreas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_QCWaterfallAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonMainMenu_Click()

    Me.Hide
    Call Mod_MainMenu.TrackerMainMenu

End Sub

Private Sub ButtonWaterfall_Click()

    boolQCWaterfallButton = True
    Me.Hide

End Sub

Private Sub Userform_Activate()
    
    'initialize color termination points
    G1Black = RGB(1, 1, 1)
    G2Red = RGB(192, 0, 0)
    G3Black = RGB(3, 3, 3) 'Again, do not waterfall G3
    G4Greens = RGB(0, 102, 0)
    G5Black = RGB(5, 5, 5)
    G6Black = RGB(6, 6, 6) 'ignore this section
    G7Black = RGB(7, 7, 7) 'ignore this section
    G8Red = RGB(255, 5, 5)
    
    'QC start line column
    G1StartLine = 0
    G2StartLine = 0
    G3StartLine = 0
    G4StartLine = 0
    G5StartLine = 0
    G6StartLine = 0
    G7StartLine = 0
    G8StartLine = 0
    
    'QC termination line column
    G1EndLine = 0
    G2EndLine = 0
    G3EndLine = 0
    G4EndLine = 0
    G5EndLine = 0
    G6EndLine = 0
    G7EndLine = 0
    G8EndLine = 0
    
    'QC booleans
    G1bool = False
    G2bool = False
    G3bool = False
    G4bool = False
    G5bool = False
    G6bool = False
    G7bool = False
    G8bool = False

End Sub

Private Sub CheckBox1_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox1.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G1bool = True
        
        'find start of group (first group so no need to search)
        G1StartLine = 3
        
        'find end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'end line found
            If gcll.Interior.Color = G1Black Then
                G1EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G1EndLine - G1StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox1.Value = False Then
        
        'group boolean
        G1bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G1EndLine - G1StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox2_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox2.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G2bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G1Black Then
                G2StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G2Red Then
                G2EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G2EndLine - G2StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox2.Value = False Then

        'group boolean
        G2bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G2EndLine - G2StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox3_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox3.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G3bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G2Red Then
                G3StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G3Black Then
                G3EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G3EndLine - G3StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox3.Value = False Then

        'group boolean
        G3bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G3EndLine - G3StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox4_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox4.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G4bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G3Black Then
                G4StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G4Greens And gcll.Offset(0, 2).Interior.Color = G4Greens Then
                G4EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G4EndLine - G4StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox4.Value = False Then

        'group boolean
        G4bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G4EndLine - G4StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox5_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox5.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G5bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            If gcll.Column >= 3 Then
                'start line found
                If gcll.Interior.Color = G4Greens And gcll.Offset(0, -2).Interior.Color = G4Greens Then
                    G5StartLine = gcll.Column + 1
                'end line found
                ElseIf gcll.Interior.Color = G5Black Then
                    G5EndLine = gcll.Column - 1
                    Exit For
                End If
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G5EndLine - G5StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox5.Value = False Then

        'group boolean
        G5bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G5EndLine - G5StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox6_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox6.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G6bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G5Black Then
                G6StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G6Black Then
                G6EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G6EndLine - G6StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox6.Value = False Then

        'group boolean
        G6bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G6EndLine - G6StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox7_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox7.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G7bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G6Black Then
                G7StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G7Black Then
                G7EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G7EndLine - G7StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox7.Value = False Then

        'group boolean
        G7bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G7EndLine - G7StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub

Private Sub CheckBox8_Click()

    Dim gcll As Range
    
    'selected
    If CheckBox8.Value = True Then
        
        'enable waterfall button
        ButtonWaterfall.Enabled = True
        
        'group boolean
        G8bool = True
        
        'find start and end of group
        For Each gcll In Worksheets("Quality Clinic").Range("1:1")
            'start line found
            If gcll.Interior.Color = G7Black Then
                G8StartLine = gcll.Column + 1
            'end line found
            ElseIf gcll.Interior.Color = G8Red Then
                G8EndLine = gcll.Column - 1
                Exit For
            End If
        Next gcll
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) + ((G8EndLine - G8StartLine) + 1) * 5)
        
        
    'not selected
    ElseIf CheckBox8.Value = False Then

        'group boolean
        G8bool = False
        
        'Estimated time
        Label2.Caption = Str(Int(Label2.Caption) - ((G8EndLine - G8StartLine) + 1) * 5)
        
    End If
    
    'disable waterfall button if nothing is selected
    If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False And CheckBox4.Value = False And CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox8.Value = False Then
        ButtonWaterfall.Enabled = False
    End If

End Sub
