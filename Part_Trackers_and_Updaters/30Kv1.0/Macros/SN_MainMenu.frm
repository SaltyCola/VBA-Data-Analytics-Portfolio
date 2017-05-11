VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SN_MainMenu 
   Caption         =   "Tracker Updater Program"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "SN_MainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SN_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonAsBuilt_Click()

    Me.Hide
    Call Mod_WIPUpdater.SNAsBuilt

End Sub

Private Sub ButtonEngSetHandler_Click()

    Me.Hide
    'delete black columns first
    Call Mod_WIPUpdater.DeleteBlackColumns
    'finalize engine sets
    Call Mod_WIPUpdater.EngineSetHandler

End Sub

Private Sub ButtonLaunchSN_Click()
    
    Me.Hide
    Call Mod_WIPUpdater.SNNewSN
    
End Sub

Private Sub ButtonQC_Click()
'NOTHING YET
End Sub

Private Sub ButtonQCtoWIP_Click()

    Me.Hide
    Call Mod_WIPUpdater.DeleteBlackColumns
    Call Mod_WIPUpdater.SNQCtoWIP

End Sub

Private Sub ButtonShipped_Click()

    Me.Hide
    Call Mod_WIPUpdater.SNShipped

End Sub

Private Sub ButtonSlow_Click()

    Me.Hide
    Call Mod_WIPUpdater.SNSlowPartsAnalyzer

End Sub

Private Sub ButtonWaterfallTracker_Click()
    
    'hide menu
    Me.Hide
    
    'WIP or QC form
    wtrfll_WIPorQC = 0
    Set frmWtrfllWIPorQC = New SN_WtrfllWIPorQC
    frmWtrfllWIPorQC.Show
    
    
    If wtrfll_WIPorQC = 2 Then
    '================================================================================================='
    '========================================Waterfalling QC=========================================='
        Dim g As Double
        Dim rngWtrFllQC As Range
        Dim cllqc As Range
        Dim startCllAd As String
        Dim endCllAd As String
        
        'initialize variables
        boolWtrFllDTGFirstRun = True
        snCount = 0
        secs = 0
        mins = 0
        hrs = 0
        
        'initialize booleans
        boolQCWaterfallButton = False
        G1bool = False
        G2bool = False
        G3bool = False
        G4bool = False
        G5bool = False
        G6bool = False
        G7bool = False
        G8bool = False
        
        'QC waterfall areas form
        Set frmQCWaterfallAreas = New SN_QCWaterfallAreas
        frmQCWaterfallAreas.Show
        
        'exit sub if waterfall button was not hit
        If Not boolQCWaterfallButton Then
            GoTo lineExitSub
        End If
        
        'boolWtrFllDTG
        boolWtrFllDTG = True
        
        'load duplicate to group form
        Set frmWtrFllDTG = New SN_DuplicateToGroup
        
        'disable controls
        frmWtrFllDTG.Caption = "Waterfall QC tab"
        frmWtrFllDTG.Label1.Caption = "Quality Clinic SN's"
        frmWtrFllDTG.TextBox.Enabled = False
        frmWtrFllDTG.TextBox.Visible = False
        frmWtrFllDTG.ListBox1.Locked = True
        frmWtrFllDTG.ConfirmButton.Caption = "Waterfall QC"
        frmWtrFllDTG.ConfirmButton.Locked = True
        frmWtrFllDTG.CancelButton.Enabled = False

        'show form
        frmWtrFllDTG.Show vbModeless
        DoEvents

        'iterate groups
        For g = 1 To 8
        
            'group 1
            If g = 1 And G1bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G1StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G1EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G1Black Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G1Black) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 2
            ElseIf g = 2 And G2bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G2StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G2EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G2Red Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G2Red) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 3
            ElseIf g = 3 And G3bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G3StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G3EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G3Black Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G3Black) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 4
            ElseIf g = 4 And G4bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G4StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G4EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G4Greens Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G4Greens) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 5
            ElseIf g = 5 And G5bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G5StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G5EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G5Black Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G5Black) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 6
            ElseIf g = 6 And G6bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G6StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G6EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G6Black Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G6Black) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 7
            ElseIf g = 7 And G7bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G7StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G7EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G7Black Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G7Black) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            'group 8
            ElseIf g = 8 And G8bool Then
                'start and end cells for range
                startCllAd = Worksheets("Quality Clinic").Cells(6, G8StartLine).Address
                endCllAd = Worksheets("Quality Clinic").Cells(6, G8EndLine).Address
                'iterate SN's
                Set rngWtrFllQC = Worksheets("Quality Clinic").Range(startCllAd & ":" & endCllAd)
                For Each cllqc In rngWtrFllQC
                    'show progress
                    Application.Goto cllqc, Scroll:=True
                    'exit iteration at end line
                    If cllqc.Interior.Color = G8Red Then: Exit For
                    'if SN, populate listbox
                    If (cllqc.Column > 2) And Not (cllqc.Value = "") And Not (cllqc.Interior.Color = G8Red) Then
                        frmWtrFllDTG.ListBox1.AddItem (cllqc.Value)
                        'scroll to newest item
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                        frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
                    End If
                Next cllqc
                
            End If
        Next g
        
        'assign list box length to sncount
        snCount = frmWtrFllDTG.ListBox1.ListCount
        MsgBox snCount
        
        'ulock confirm button
        frmWtrFllDTG.ConfirmButton.Locked = False
        frmWtrFllDTG.ConfirmButton.Enabled = True
        
        'scroll to beginning
        Application.Goto Worksheets("Quality Clinic").Range("C6"), Scroll:=True
    '========================================Waterfalling QC=========================================='
    '================================================================================================='
    
    
    ElseIf wtrfll_WIPorQC = 1 Then
    '================================================================================================='
    '========================================Waterfalling WIP========================================='
        Dim rngWtrFllTracker As Range
        Dim cll As Range
        
        'define variables
        Set rngWtrFllTracker = Worksheets("NEO 5322121").Range("6:6")
        boolWtrFllDTGFirstRun = True
        snCount = 0
        secs = 0
        mins = 0
        hrs = 0
        
        '====================Confimation Message with estimated time to complete=====================
        'initialize boolean
        boolWtrFllCont = False
        
        'count serial numbers
        For Each cll In rngWtrFllTracker
            If cll.Interior.Color = RGB(255, 0, 0) Then
                Exit For
            ElseIf (cll.Column > 2) And Not (cll.Interior.Color = RGB(0, 0, 0)) Then
                snCount = snCount + 1
            End If
        Next cll
        
        'load ETA Message
        Set frmWtrFllETA = New MsgWaterfallETA
        
        'seconds:
        secs = snCount * 5
        
        'minutes: if ETA (seconds) has no remainder after division by 60
        If (secs / 60 >= 1) Then
            mins = ((secs - (secs Mod 60)) / 60)
            secs = (secs Mod 60)
        
        'hours: if ETA (seconds) has no remainder after division by 3600
            If (mins / 60 >= 1) Then
                hrs = ((mins - (mins Mod 60)) / 60)
                mins = (mins Mod 60)
            End If
        End If
        
        'fill in ETA Label
        If hrs > 0 Then
            frmWtrFllETA.Label3.Caption = Str(hrs) & " hrs, " & Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins > 0) Then
            frmWtrFllETA.Label3.Caption = Str(mins) & " mins, " & Str(secs) & " secs"
        ElseIf (hrs = 0) And (mins = 0) Then
            frmWtrFllETA.Label3.Caption = Str(secs) & " secs"
        End If
        
        'show continue message
        frmWtrFllETA.Show
        
        'continue?
        If Not boolWtrFllCont Then
            Exit Sub
        End If
        '===========================================================================================
        
        'boolWtrFllDTG
        boolWtrFllDTG = True
        
        'call black line deleter
        Call DeleteBlackColumns
        
        'load duplicate to group form
        Set frmWtrFllDTG = New SN_DuplicateToGroup
        
        'disable controls
        frmWtrFllDTG.Caption = "Waterfall WIP tab"
        frmWtrFllDTG.Label1.Caption = "Tracker SN's"
        frmWtrFllDTG.TextBox.Enabled = False
        frmWtrFllDTG.TextBox.Visible = False
        frmWtrFllDTG.ListBox1.Locked = True
        frmWtrFllDTG.ConfirmButton.Caption = "Waterfall"
        frmWtrFllDTG.ConfirmButton.Locked = True
        frmWtrFllDTG.CancelButton.Enabled = False
        
        'show form
        frmWtrFllDTG.Show vbModeless
        DoEvents
        
        'iterate SN's
        For Each cll In rngWtrFllTracker
        
            'show progress
            Application.Goto cll, Scroll:=True
            
            'exit iteration at red line
            If cll.Interior.Color = RGB(255, 0, 0) Then: Exit For
        
            'if SN, populate listbox
            If (cll.Column > 2) And Not (cll.Value = "") And Not (cll.Interior.Color = RGB(0, 0, 0)) Then
                frmWtrFllDTG.ListBox1.AddItem (cll.Value)
                'scroll to newest item
                frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = True
                frmWtrFllDTG.ListBox1.Selected(frmWtrFllDTG.ListBox1.ListCount - 1) = False
            End If
        
        Next cll
        
        'ulock confirm button
        frmWtrFllDTG.ConfirmButton.Locked = False
        frmWtrFllDTG.ConfirmButton.Enabled = True
        
        'scroll to beginning
        Application.Goto Worksheets("NEO 5322121").Range("C6"), Scroll:=True
    '========================================Waterfalling WIP========================================='
    '================================================================================================='
    End If

lineExitSub:
    'reset WIP or QC
    wtrfll_WIPorQC = 0

End Sub

Private Sub ButtonWIP_Click()

    Me.Hide
    Call Mod_WIPUpdater.DeleteBlackColumns
    Call Mod_WIPUpdater.SNSearchBox

End Sub

Private Sub ButtonWIPtoQC_Click()
'NOTHING YET
End Sub

Private Sub SN_MainMenu_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Unload (Me)
    End If

End Sub
