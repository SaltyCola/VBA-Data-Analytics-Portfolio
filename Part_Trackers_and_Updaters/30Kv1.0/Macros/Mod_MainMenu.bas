Attribute VB_Name = "Mod_MainMenu"
Option Explicit

Public frmSNMainMenu As SN_MainMenu
Public frmMsgSavingBackupFile As MsgSavingBackupFile
Public frmMsgTempDelBL As MsgTempDelBL
Public frmMsgSeparatingEngSets As MsgSeparatingEngSets
Public frmSNSearch As SN_Search
Public frmMsgWaterfall As MsgWaterfall
Public frmSNMatchList As SN_MatchList
Public frmSNInfoPage As SN_InfoPage
Public frmSNQCPartInfo As SN_QCPartInfo
Public frmSNDeleteError As SN_DeleteError
Public frmSNDeleteErrorDTG As SN_DeleteError
Public frmSNDeleteErrorWtrFllTrkr As SN_DeleteError
Public frmSNDeleteErrorSNIPCI As SN_DeleteError
Public frmSNDuplicate As SN_DuplicateToGroup
Public frmWtrFllDTG As SN_DuplicateToGroup
Public frmWtrFllETA As MsgWaterfallETA
Public frmSNAsBuilt As SN_AsBuilt
Public frmSNSlowParts As SN_SlowParts
Public frmSNNewSN As SN_NewSN
Public frmSNShipped As SN_Shipped
Public frmQCtoWIP As SN_QCtoWIP
Public frmWtrfllWIPorQC As SN_WtrfllWIPorQC
Public frmQCWaterfallAreas As SN_QCWaterfallAreas

Public snTitleCell As Range
Public snSearchTxt As String
Public snSearchCol As Long
Public boolCanceled As Boolean
Public intError As Boolean
Public boolTog As Boolean
Public boolTogsAllowed As Boolean
Public boolRTOTog As Boolean
Public txtBoxTxt As String
Public txtBoxClr As Long
Public togClr As Long
Public TodaysDate As String
Public SNIPRow As Double
Public SNIPTog As String
Public SNIPSearchCurrentRow As Double
Public SNIPLastDate As String
Public clrBlank As Long
Public SNMLArray() As String
Public SNMLArrayCnt As Double
Public intSNMLType As Double

Public arraySNBackupVal() As String
Public arraySNBackupClr() As Long
Public arraySNUpdatedVal() As String
Public arraySNUpdatedClr() As Long
Public arrayUpdatesOnlyVal() As String
Public arrayUpdatesOnlyClr() As Long
Public arrayWaterfallVal() As String
Public arrayWaterfallClr() As Long
Public arraySNTemplateVal() As String
Public arraySNTemplateClr() As Long

Public rngErrorCell As String
Public boolFlagCell As Boolean
Public boolFlagCellClear As Boolean
Public boolSNDTGFirstRun As Boolean
Public boolSNDTGAdd As Boolean
Public SNDTGListedSNCell As Range
Public boolSNDTGsnNotFound As Boolean
Public boolWaterfallingTrackerMsgBox As Boolean
Public boolWtrFllDTG As Boolean
Public boolWtrFllDTGFirstRun As Boolean
Public boolWtrFllCont As Boolean
Public boolAsBuilt As Boolean
Public boolDoNotClose As Boolean
Public boolMaintoAB As Boolean
Public boolSNMLtoAB As Boolean
Public abCol As Double
Public boolSNMLtoShipped As Boolean
Public boolShippedDoNotAdd As Boolean
Public CurrentEngineSet As Double
Public shpdFinalBlackLine As Double

Public snCount As Double
Public hrs As Double
Public mins As Double
Public secs As Double

Public strNewSN As String
Public redlineRng As Range
Public redlineCell As Range
Public redlineInt As Double
Public redlineColAddress As String
Public replaceInt As Double
Public wtrfll_WIPorQC As Double
Public boolQCWaterfallButton As Boolean

'public QC color termination points
Public G1Black As Long
Public G2Red As Long
Public G3Black As Long 'ignore this section (Do not waterfall)
Public G4Greens As Long
Public G5Black As Long
Public G6Black As Long 'ignore this section
Public G7Black As Long 'ignore this section
Public G8Red As Long
'QC start column numbers
Public G1StartLine As Double
Public G2StartLine As Double
Public G3StartLine As Double
Public G4StartLine As Double
Public G5StartLine As Double
Public G6StartLine As Double
Public G7StartLine As Double
Public G8StartLine As Double
'QC termination column numbers
Public G1EndLine As Double
Public G2EndLine As Double
Public G3EndLine As Double
Public G4EndLine As Double
Public G5EndLine As Double
Public G6EndLine As Double
Public G7EndLine As Double
Public G8EndLine As Double
'QC group booleans to know what to waterfall
Public G1bool As Boolean
Public G2bool As Boolean
Public G3bool As Boolean
Public G4bool As Boolean
Public G5bool As Boolean
Public G6bool As Boolean
Public G7bool As Boolean
Public G8bool As Boolean

Public Sub SaveFileBackup()
'saves a backup file in backup directory within current directory

    Dim MyPath As String
    Dim MyNow As String
    Dim nowStr As String
    Dim MyBackupPath As String
    
    'define variables
    MyPath = ActiveWorkbook.Path
    MyNow = Str(Now)
    nowStr = Replace(Replace(MyNow, ":", "."), "/", "-")
    
    'show Saving Backup Message
    Set frmMsgSavingBackupFile = New MsgSavingBackupFile
    frmMsgSavingBackupFile.Show vbModeless
    DoEvents
    
    'deactivate alerts
    Application.DisplayAlerts = False
    
    'delete previous backup folders
    Dim i As Double
    Dim Fs As Object
    For i = 1 To 7
        If Not Len(Dir(MyPath & "\BACKUPS - 30K Update Program " & Replace(Str(DateAdd("d", -(i), Date)), "/", "-"), vbDirectory)) = 0 Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.DeleteFolder (MyPath & "\BACKUPS - 30K Update Program " & Replace(Str(DateAdd("d", -(i), Date)), "/", "-")), True
        End If
    Next i
    
    'create backup folder if DNE
    If Len(Dir(MyPath & "\BACKUPS - 30K Update Program " & Replace(Str(Date), "/", "-"), vbDirectory)) = 0 Then
        MkDir (MyPath & "\BACKUPS - 30K Update Program " & Replace(Str(Date), "/", "-"))
    End If
    
    'define backup path
    MyBackupPath = (MyPath & "\BACKUPS - 30K Update Program " & Replace(Str(Date), "/", "-") & "\")
    
    'save backup copy
    ActiveWorkbook.SaveCopyAs Filename:=MyBackupPath & " (" & nowStr & ") " & ActiveWorkbook.Name
    
    'save current file
    ActiveWorkbook.Save
    
    'reactivate alerts
    Application.DisplayAlerts = True
    
    'close Saving Backup Message
    frmMsgSavingBackupFile.Hide

    
    'Call Main Menu!
    Call TrackerMainMenu
    

End Sub

Public Sub TrackerMainMenu()
'Tracker Updater Main Menu

    'Initialize variables
    boolCanceled = False
    intError = False
    boolTog = False
    boolTogsAllowed = False
    boolRTOTog = False
    boolSNDTGsnNotFound = False
    txtBoxTxt = ""
    txtBoxClr = 0
    togClr = 0
    TodaysDate = Str(Date)
    SNIPRow = 0
    SNIPTog = "0"
    SNIPSearchCurrentRow = 44
    ReDim SNMLArray(0)
    SNMLArrayCnt = 0
    intSNMLType = 0
    clrBlank = RGB(255, 255, 255)
    boolAsBuilt = False
    boolMaintoAB = True
    boolSNMLtoAB = False
    abCol = 0
    boolSNMLtoShipped = False
    boolShippedDoNotAdd = False
    CurrentEngineSet = 0
    shpdFinalBlackLine = 0
    boolWtrFllDTGFirstRun = True
    boolDoNotClose = False
    
    'initialize arrays
    ReDim arraySNBackupVal(56)
    ReDim arraySNBackupClr(56)
    ReDim arraySNUpdatedVal(56)
    ReDim arraySNUpdatedClr(56)
    ReDim arrayUpdatesOnlyVal(56)
    ReDim arrayUpdatesOnlyClr(56)
        '56th entry is SN column before waterfall!!!!
    ReDim arrayWaterfallVal(57)
        '============================================
    ReDim arrayWaterfallClr(56)
    
    '=================Updates Only Arrays Initializer================='
    Dim arr As Double
    For arr = 1 To 56
        arrayUpdatesOnlyVal(arr) = "!===N/A===!"
        arrayUpdatesOnlyClr(arr) = 102030405
    Next arr
    '=================Updates Only Arrays Initializer================='

    'initialize booleans
    boolWtrFllDTG = False

    'Load SN_MainMenu form
    Set frmSNMainMenu = New SN_MainMenu
    
    'Uncomment to Disable buttons under construction=============================
    
    'frmSNMainMenu.ButtonLaunchSN.Enabled = False
    'frmSNMainMenu.ButtonShipped.Enabled = False
    'frmSNMainMenu.ButtonWIP.Enabled = False
    frmSNMainMenu.ButtonQC.Enabled = False
    frmSNMainMenu.ButtonWIPtoQC.Enabled = False
    'frmSNMainMenu.ButtonQCtoWIP.Enabled = False
    'frmSNMainMenu.ButtonSlow.Enabled = False
    'frmSNMainMenu.ButtonAsBuilt.Enabled = False
    'frmSNMainMenu.ButtonEngSetHandler.Enabled = False
    'frmSNMainMenu.ButtonWaterfallTracker.Enabled = False
    
    '============================================================================
    
    'show SN_MainMenu
    frmSNMainMenu.Show

End Sub
