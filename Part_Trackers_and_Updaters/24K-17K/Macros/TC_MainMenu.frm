VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TC_MainMenu 
   Caption         =   "Tracker Updater Program"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "TC_MainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TC_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public bWaterfalledAfterChanges As Boolean 'boolean to track whether most recent changes applied by waterfaller

Private Sub ButtonWIP_Click()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
        'none required
    
    'hide Main Menu
    Me.Hide
    
    'call Sub Director
    Call SubDirector
    
    'set waterfall boolean
    Me.bWaterfalledAfterChanges = False
    
    'show UC Search (without tracker movement)
    ufUCSearch.Show

End Sub

Private Sub ButtonWaterfallTracker_Click()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
    bWaterfallSort = True
    bCreateWIPSheetCopy = True
    bClearWIP = True
    bWriteWIP = True
    bDeleteWIPSheetCopy = True
    
    'hide Main Menu
    Me.Hide
    
    'call Sub Director
    Call SubDirector
    
    'set waterfall boolean
    Me.bWaterfalledAfterChanges = True
    
    'show Main Menu (without tracker movement)
    Me.Show

End Sub

Private Sub ButtonTAI_Click()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
        'none required
    
    'hide main menu
    Me.Hide
    
    'Update TAI Status from IRO Log
    Call IROUpdate
    
    'Reread WIP in preparation for any more updates
    Call ReadWIP
    
    'show Main Menu (without tracker movement)
    Me.Show

End Sub

Private Sub UserForm_Initialize()

    'initialize waterfall boolean
    Me.bWaterfalledAfterChanges = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Unload everything from memory.

    If CloseMode = 0 And Me.bWaterfalledAfterChanges Then
        End
    ElseIf CloseMode = 0 And Not Me.bWaterfalledAfterChanges Then
        'prevent closing userform
        Cancel = True
        'Warning Message
        MsgBox "If you close the program before waterfalling, any updates you have made will not be saved. Click the exit button again to close without updates, or click the waterfall button to apply updates.", , "WARNING"
        'Change waterfalled after changes boolean back to true to allow exiting without updates
        Me.bWaterfalledAfterChanges = True
    End If

End Sub

