Attribute VB_Name = "TC_TrackerControl"
'To Alter the code to work for other part numbers,
'and different tracker organizations search for: <==Alter
'Once completed, change <==Alter to <==Altered

'Also, must create new TC_CellColor Class and TC_ColorChoice for each new part number


'If a color is added then the CellColor Class and Color Choice UserForm must be updated accordingly.
'Any rows added must be updated within this module and within UCDisplay.


'Declarations
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

'UserForms
Public ufMainMenu As TC_MainMenu 'Tracker Updater Program Main Menu
Public ufUCSearch As TC_UCSearch 'Unit Column search / group search window
Public ufColorChooser As TC_ColorChoice_24K 'Allows user to select cell colors from a predetermined list
Public ufUCDisplay As TC_UCDisplay 'Userform version of a Unit Column, allows live editing in a controlled environment
Public ufYesNoMsg As TC_YesNoMsg 'Gives user a second chance to back out of a button press in case of accidental button pressing
Public lBar As TC_LoadingBar 'Loading Bar that will be used throughout program
Public ufCellReadError As TC_CellError 'Allows user to edit cell value errors at runtime

'Event Handler Collections
Public cEH_UCDisplay As Collection 'Collection of Event Handler Targets with same actions
Public cEH_ColorChoice As Collection 'Collection of Event Handler Targets with same actions

'Public Variables
Public BookWIP As Workbook 'Workbook containing WIP
Public SheetWIP As Worksheet 'Worksheet containing WIP
Public SheetWIP_BACKUP As Worksheet 'Worksheet containing WIP Sheet Backup
Public SheetShipped As Worksheet 'Worksheet containing all shipped UC's
Public SheetSummary24K As Worksheet 'Worksheet containing 24K WIP Summary
Public SheetSummarySK As Worksheet 'Worksheet containing SK WIP Summary
Public SheetCharts As Worksheet 'Worksheet on which the charts are located
Public ArrWIP() As TC_UnitColumn 'Array to store all Unit Columns in WIP '<==Alter (if more than one part number, i.e. SGVs, must erase ArrWIP at the start of each new part number before ReadWIP)
Public ArrSummary() As TC_SummaryColumn 'Array to store all Summary Columns for writing to the Summary page
Public ArrListBox() As TC_UnitColumn 'Unit Column array for keeping track of all current UCSearch listbox entries
Public StartWIP As Range 'first cell of WIP Range 1 (top left corner)
Public EndWIP As Range 'last cell of WIP Range 1 (bottom right corner)
Public StartWIP2 As Range 'first cell of WIP Range 2 (top left corner)
Public EndWIP2 As Range 'last cell of WIP Range 2 (bottom right corner)
Public StartWIP3 As Range 'first cell of WIP Range 3 (top left corner)
Public EndWIP3 As Range 'last cell of WIP Range 3 (bottom right corner)
Public StartWIP4 As Range 'first cell of WIP Range 4 (top left corner)
Public EndWIP4 As Range 'last cell of WIP Range 4 (bottom right corner)
'Subroutine Booleans
    Public bSaveBackupFile As Boolean 'Boolean to tell Program to call this subroutine
    Public bReadWIP As Boolean 'Boolean to tell Program to call this subroutine
    Public bInitializeEventHandlerCollections As Boolean 'Boolean to tell Program to call this subroutine
    Public bInitializeUserForms As Boolean 'Boolean to tell Program to call this subroutine
    Public bCompleteSummary As Boolean 'Boolean to tell Program to call this subroutine
    Public bWaterfallSort As Boolean 'Boolean to tell Program to call this subroutine
    Public bCreateWIPSheetCopy As Boolean 'Boolean to tell Program to call this subroutine
    Public bClearWIP As Boolean 'Boolean to tell Program to call this subroutine
    Public bWriteWIP As Boolean 'Boolean to tell Program to call this subroutine
    Public bDeleteWIPSheetCopy As Boolean 'Boolean to tell Program to call this subroutine

'24K Tracker Specific Variables
Public arrNonDateOpRows(1 To 3) As Integer 'Array that holds the row numbers of the Op Rows that can have non-date values

Public Sub TrackerUpdaterProgram()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
    bSaveBackupFile = True
    bReadWIP = True
    bInitializeEventHandlerCollections = True
    bInitializeUserForms = True
    
    'call T.U.P.
    Call TUP_Initialize
    
    'show Main Menu (without tracker movement)
    ufMainMenu.Show

End Sub

Public Sub SummaryUpdaterTUP()

    'initialize booleans
    Call InitializePublicBooleans
    
    'set booleans
    bReadWIP = True
    bInitializeEventHandlerCollections = True
    bInitializeUserForms = True
    bCompleteSummary = True
    
    'call T.U.P.
    Call TUP_Initialize

End Sub

Public Sub InitializePublicBooleans()

    'initialize booleans
    bSaveBackupFile = False
    bReadWIP = False
    bInitializeEventHandlerCollections = False
    bInitializeUserForms = False
    bCompleteSummary = False
    bWaterfallSort = False
    bCreateWIPSheetCopy = False
    bClearWIP = False
    bWriteWIP = False
    bDeleteWIPSheetCopy = False

End Sub

Public Sub TUP_Initialize()

    Dim i As Integer 'iterator
    Dim cRedLine As Range 'range iterator for finding EndWIP range
    Dim rowsWIP As Integer 'integer that must be manually set to hold the final row in the WIP grid
    Dim b1stVisible As Boolean 'boolean for grabbing StartWIP position based on first unhidden Unit Column
    Dim cntRedLines As Integer 'Counter for grabbing starting and ending ranges for each of the 4 WIP sections
    
    'initialize public variables
    Set BookWIP = ActiveWorkbook '<==Alter
    Set SheetWIP = BookWIP.Worksheets("24K tab") '<==Alter
    Set SheetShipped = BookWIP.Worksheets("Shipped Sets 24K") '<==Alter
    Set SheetSummary24K = BookWIP.Worksheets("Tracker Summaries 24K") '<==Alter
    Set SheetSummarySK = BookWIP.Worksheets("Tracker Summaries SK 24K") '<==Alter
    Set SheetCharts = BookWIP.Worksheets("Charts 24K") '<==Alter
        
    'initialize private variables
    rowsWIP = 48 '<==Alter
    b1stVisible = True
    cntRedLines = 1
    
    'initialize loading bar
    Set lBar = New TC_LoadingBar
    
    'initialize cell value reading error corrector
    Set ufCellReadError = New TC_CellError
    
    'initialize yes or no choice userform
    Set ufYesNoMsg = New TC_YesNoMsg
        
    'activate WIP worksheet
    SheetWIP.Activate
    
    'initialize startWIP and EndWIP variables
    For Each cRedLine In Range("1:1")
        'first visible unit column found
        If (b1stVisible) And (cRedLine.Column > 4) And Not (cRedLine.EntireColumn.Hidden) Then '<==Alter
            Set StartWIP = cRedLine
            'change boolean to stop looking for StartWIP section 1 position
            b1stVisible = False
        'first redline found
        ElseIf cntRedLines = 1 And cRedLine.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            'assign end of WIP section 1
            Set EndWIP = Cells(rowsWIP, (cRedLine.Column - 1))
            'assign start of WIP section 2
            Set StartWIP2 = Cells(1, (cRedLine.Column + 1))
            'Set up to find next section variables
            cntRedLines = 2
        'second redline found
        ElseIf cntRedLines = 2 And cRedLine.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            'assign end of WIP section 2
            Set EndWIP2 = Cells(rowsWIP, (cRedLine.Column - 1))
            'assign start of WIP section 3
            Set StartWIP3 = Cells(1, (cRedLine.Column + 1))
            'Set up to find next section variables
            cntRedLines = 3
        'third redline found
        ElseIf cntRedLines = 3 And cRedLine.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            'assign end of WIP section 3
            Set EndWIP3 = Cells(rowsWIP, (cRedLine.Column - 1))
            'assign start of WIP section 4
            Set StartWIP4 = Cells(1, (cRedLine.Column + 1))
            'Set up to find next section variables
            cntRedLines = 4
        'fourth redline found
        ElseIf cntRedLines = 4 And cRedLine.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            'assign end of WIP section 4
            Set EndWIP4 = Cells(rowsWIP, (cRedLine.Column - 1))
            'end for loop
            Exit For
        End If
    Next cRedLine
    
    'set 24K Tracker Specific Variables
    For i = 1 To 3
        arrNonDateOpRows(i) = i + 6
    Next i
    
    'call sub director
    Call SubDirector

End Sub

Public Sub SaveBackupFile()
'Saves a backup file in backup directory in personal computer documents.
'Also removes all backups starting yesterday and going back for a month.
'This should ensure the backups do not take up any more space than they need to.

    Dim ufMsgSaveBackup As TC_MsgSaveBackup 'Saving Message userform
    Dim myPath As String 'path string to "my documents"
    Dim myNow As String 'string of current date-time
    Dim nowStr As String 'altered MyNow string for valid use within a path
    Dim myBackupPath As String 'Final path for backup file saving
    Dim i As Integer 'iterator for deleting previous backup folders
    Dim objFolder As Object 'object variable representing folders to delete
    
    'define variables
    myPath = Environ$("USERPROFILE") & "\Documents" 'grab local documents folder path
    myNow = Str(Now)
    nowStr = Replace(Replace(myNow, ":", "."), "/", "-")
    
    'show Saving Backup Message
    Set ufMsgSaveBackup = New TC_MsgSaveBackup
    ufMsgSaveBackup.Show vbModeless
    DoEvents
    
    'deactivate alerts
    Application.DisplayAlerts = False
    
    'delete previous month's backup folders
    For i = 1 To 30
        If Not Len(Dir(myPath & "\BACKUPS - 24K Tracker\" & Replace(Str(DateAdd("d", -(i), Date)), "/", "-"), vbDirectory)) = 0 Then '<==Alter
            Set objFolder = CreateObject("Scripting.FileSystemObject")
            objFolder.DeleteFolder (myPath & "\BACKUPS - 24K Tracker\" & Replace(Str(DateAdd("d", -(i), Date)), "/", "-")), True '<==Alter
        End If
    Next i
    
    'create backup folder if DNE
    If Len(Dir(myPath & "\BACKUPS - 24K Tracker\", vbDirectory)) = 0 Then '<==Alter
        MkDir (myPath & "\BACKUPS - 24K Tracker\") '<==Alter
    End If
    
    'create Date folder if DNE
    If Len(Dir(myPath & "\BACKUPS - 24K Tracker\" & Replace(Str(Date), "/", "-") & "\", vbDirectory)) = 0 Then '<==Alter
        MkDir (myPath & "\BACKUPS - 24K Tracker\" & Replace(Str(Date), "/", "-") & "\") '<==Alter
    End If
    
    'define backup path
    myBackupPath = (myPath & "\BACKUPS - 24K Tracker\" & Replace(Str(Date), "/", "-") & "\") '<==Alter
    
    'save backup copy
    BookWIP.SaveCopyAs Filename:=myBackupPath & "(" & nowStr & ") " & BookWIP.Name
    
'TEMPORARY CODE===============================================================================
Dim bSaveMainFile As Boolean
'Ask before main file save, because of current internet connection issues
ufYesNoMsg.YesNoMsgInitialize ("Would you like to save this file?" & vbNewLine & "(A backup file was still saved to 'My Computer')")
bSaveMainFile = ufYesNoMsg.bYesNoMsg
'TEMPORARY CODE===============================================================================
'Also Delete if statement below when removing this temporary code
'BUT LEAVE BookWIP.Save
    
    'save current file
    If bSaveMainFile Then: BookWIP.Save
    
    'reactivate alerts
    Application.DisplayAlerts = True
    
    'close and delete Saving Backup Message
    ufMsgSaveBackup.Hide
    Set ufMsgSaveBackup = Nothing

End Sub

Public Sub SubDirector()

    
    'call backup saver (saves to personal hard drive)
    If bSaveBackupFile Then: Call SaveBackupFile
    
    
    'all subs controlled by public booleans set in starting sub
    
    If bReadWIP Then: Call ReadWIP 'requires loading bar
    
    If bInitializeEventHandlerCollections Then: Call InitializeEventHandlerCollections 'initialize event handlers
    If bInitializeUserForms Then: Call InitializeUserForms 'initialize userforms
        'reset these two booleans to false for the remainder of the program
        bInitializeEventHandlerCollections = False
        bInitializeUserForms = False
    
    If bCompleteSummary Then: Call CompleteSummary 'requires loading bar
    
    If bWaterfallSort Then: Call WaterfallSort 'requires loading bar
    
    If bCreateWIPSheetCopy Then: Call CreateWIPSheetCopy
    
    If bClearWIP Then: Call ClearWIP
    
    If bWriteWIP Then: Call WriteWIP 'requires loading bar
    
    If bDeleteWIPSheetCopy Then: Call DeleteWIPSheetCopy

End Sub

Public Sub ReadWIP()

    Dim arrTemp() As TC_UnitColumn 'Temporary Array to store all Unit Columns in WIP
    Dim tUCol As TC_UnitColumn 'Temporary Unit Column object for entering data into array
    Dim tOpRow As TC_OpRow 'Temporary Op Row object to add to unit column's operations list
    Dim sWIP As Range 'StartWIP section range object
    Dim eWIP As Range 'EndWIP section range object
    Dim c As Range 'generic range iteration object
    Dim i As Integer 'generic integer object for adding Unit Column objects to WIP array
    Dim indWIPsection As Integer 'WIP section iterator (4 sections)
    
    'activate WIP worksheet
    SheetWIP.Activate
    
    'initialize variables
    Set tUCol = New TC_UnitColumn
    i = 0
    
    
    'iterate sections of WIP
    For indWIPsection = 1 To 4
    
        'set startwip and endwip variables
        If indWIPsection = 1 Then
            Set sWIP = StartWIP
            Set eWIP = EndWIP
        ElseIf indWIPsection = 2 Then
            Set sWIP = StartWIP2
            Set eWIP = EndWIP2
        ElseIf indWIPsection = 3 Then
            Set sWIP = StartWIP3
            Set eWIP = EndWIP3
        ElseIf indWIPsection = 4 Then
            Set sWIP = StartWIP4
            Set eWIP = EndWIP4
        End If
        
        'iterate WIP section
        For Each c In Range(Cells(5, sWIP.Column), Cells(5, eWIP.Column)) '<==Alter
        
            'Delete Blank Column from cutting/inserting between tabs?
            If Application.CountA(c.EntireColumn) = 0 Then
                'not a black column or redline
                If Not c.EntireColumn.Interior.Color = RGB(0, 0, 0) And Not c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                    'show column
                    Application.GoTo c, True
                    'ask user if they want the column deleted
                    ufYesNoMsg.YesNoMsgInitialize ("A blank column was detected at: " & c.Address & vbNewLine & "Would you like to delete this column?")
                End If
                'delete column
                If ufYesNoMsg.bYesNoMsg Or c.EntireColumn.Interior.Color = RGB(0, 0, 0) And Not c.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                    'move back a column so we don't delete the variable's reference
                    Set c = c.Offset(0, -1)
                    'delete column just moved back from
                    c.Offset(0, 1).EntireColumn.Delete
                    'move to location of where column just deleted used to be
                    Set c = c.Offset(0, 1)
                End If
            End If
            
            'only WIP if not hidden and not empty
            If Not c.EntireColumn.Hidden And Application.CountA(c.EntireColumn) > 0 Then
                
                'redimension array
                i = i + 1
                ReDim Preserve arrTemp(1 To i) As TC_UnitColumn
                
                'grab Unit Column property values
                tUCol.ColumnAddress = c.Address
                tUCol.ColumnNumber = c.Column
                tUCol.PartNumber = "24K" '<==Alter
                tUCol.TrackingNumber = c.Value
                tUCol.TNumAbbr = Right(c.Value, 5) '<==Alter
                tUCol.LastDateSeen = Cells(48, c.Column).Value '<==Alter
                tUCol.WSorted = False
                tUCol.FIPL = False 'initialize to false, statement below decides if true
                If c.Interior.Color = RGB(255, 255, 0) Then: tUCol.FIPL = True '<==Alter
                'initialize indexes
                tUCol.ColorOrderIndex = 0
                tUCol.WaterfallIndex = i
                'GrabData Methods
                tUCol.Headers.GrabData Cells(1, c.Column), 5 '<==Alter
                tUCol.GrabOperationsData Cells(6, c.Column), 33 '<==Alter
                tUCol.Notes.GrabData Cells(39, c.Column), 10 '<==Alter
                'fix SN Title value in Headers
                tUCol.Headers.TitlesList(c.Row) = "Tracking Number"
                'grab color order index
                tUCol.ColorOrderIndex = tUCol.OperationsList(tUCol.LastOpCompleted).UCColorOrderIndex
                
                'set WIP section property
                tUCol.WIPSectionIndex = indWIPsection
                
                'add Unit to array
                Set arrTemp(i) = tUCol
                
                'reset tUCol object
                Set tUCol = Nothing
                Set tUCol = New TC_UnitColumn
                
                'update loading bar
                lBar.UpdateLoadingBar "Reading WIP section " & Str(indWIPsection) & " of 4...", (c.Column - sWIP.Column), (eWIP.Column - sWIP.Column)
                
            End If
        
        Next c
    
    Next indWIPsection
    
    
    'assign temp array to Public WIP array
    ArrWIP = arrTemp '<==Alter (if more than one part number i.e. SGVs)

End Sub

Public Sub InitializeEventHandlerCollections()

    'initialize collections
    Set cEH_UCDisplay = New Collection
    Set cEH_ColorChoice = New Collection

End Sub

Public Sub InitializeUserForms()

    'initialize userforms
    Set ufMainMenu = New TC_MainMenu
    Set ufUCSearch = New TC_UCSearch
    Set ufUCDisplay = New TC_UCDisplay
    Set ufColorChooser = New TC_ColorChoice_24K
    'lBar is initialized in TUP_Initialize due to use in ReadWIP sub
    'ufCellReadError is initialized in TUP_Initialize due to use in ReadWIP sub
    'ufYesNoMsg is initialized in TUP_Initialize due to use in ReadWIP sub

End Sub

Public Sub CompleteSummary()

    Dim arrSummarySeed() As Variant 'Array of colors and OpRowIndexes for each Summary Column
    Dim rTitleRow As Integer 'row number of the SC title row
    Dim rFirstListRow As Integer 'row number of the first UC list row
    Dim cFirstSC As Integer 'column number of the first SC
    Dim cLastSC As Integer 'column number of the last SC
    Dim iCats As Integer 'number of categories on this Summary Page for entry into SC objects
    Dim tSC As TC_SummaryColumn 'temp Summary Column object for entry into ArrSummary
    Dim tCellColors As TC_CellColor_24K 'temp cell color list object
    Dim dToday As Date 'Today's date for grabbing slow parts (last seen 4 or more days before today)
    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    Dim i3 As Integer 'iterator
    
    'activate summary page
    SheetSummary24K.Activate
    
    'initialize summary table property variables
    Set tCellColors = New TC_CellColor_24K '<==Alter
    rTitleRow = 5 '<==Alter
    rFirstListRow = 13 '<==Alter
    cFirstSC = 4 '<==Alter
    cLastSC = 42 '<==Alter
    iCats = 3 '<==Alter
    dToday = Date
    
    'resize arrSummary
    ReDim ArrSummary(1 To ((cLastSC - cFirstSC) + 1)) As TC_SummaryColumn
    
    
    'set summary seed '<==Alter (entire list is done manually at implementation)
    ReDim arrSummarySeed(1 To (iCats + 1), 1 To ((cLastSC - cFirstSC) + 1)) As Variant
    For i = 1 To UBound(arrSummarySeed, 2) 'iterating summary columns
        'Category Colors
        arrSummarySeed(2, i) = tCellColors.Bad 'Category 1 color
        arrSummarySeed(3, i) = tCellColors.RTO 'Category 2 color
        arrSummarySeed(4, i) = tCellColors.OldComplete 'Category 3 color
        'Op Row Indexes
        If i <= 7 Then
            arrSummarySeed(1, i) = (i) 'op row index
        ElseIf i <= 12 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 1) 'op row index
        ElseIf i <= 16 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 2) 'op row index
        ElseIf i <= 20 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 3) 'op row index
        ElseIf i <= 21 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 4) 'op row index
        ElseIf i <= 23 Then '1 hidden row (+1)
            arrSummarySeed(1, i) = (i - 3) 'op row index
        ElseIf i <= 26 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 4) 'op row index
        ElseIf i <= 38 Then '1 double row (-1)
            arrSummarySeed(1, i) = (i - 5) 'op row index
        End If
        'Fix Outlier Category Colors
        If i = 8 Or i = 13 Or i = 17 Or i = 21 Or i = 24 Or i = 27 Then 'dark green categories
            arrSummarySeed(2, i) = tCellColors.Blank 'Category 1 color
            arrSummarySeed(3, i) = tCellColors.Blank 'Category 2 color
            arrSummarySeed(4, i) = tCellColors.OldOX_RCVD 'Category 3 color
        End If
    Next i
    'set summary seed '<==Alter (entire list is done manually at implementation)
    
    
    'iterate through title row to create SCs
    For Each c In Range(Cells(rTitleRow, cFirstSC), Cells(rTitleRow, cLastSC))
    
        'initialize SC object
        Set tSC = New TC_SummaryColumn
        
        'set SCol properties
        tSC.Title = c.Value
        tSC.OpIndex = arrSummarySeed(1, ((c.Column - cFirstSC) + 1))
        tSC.NumberOfCategories = iCats
        tSC.InitializeCategories
        
        'set SCat properties
        For i = 1 To tSC.NumberOfCategories
            tSC.CategoryList(i).Title = Cells((rTitleRow + i), 1).Value
            tSC.CategoryList(i).Color = arrSummarySeed((i + 1), ((c.Column - cFirstSC) + 1))
        Next i
        
        'add tSC to ArrSummary
        Set ArrSummary(((c.Column - cFirstSC) + 1)) = tSC
    
    Next c
    
    
    'Set UC Lists for each SC's categories
    For i1 = 1 To UBound(ArrWIP)
    
        'iterate arrSummary
        For i2 = 1 To UBound(ArrSummary)
        
            'Correct Summary Column found
            If ArrWIP(i1).LastOpCompleted = ArrSummary(i2).OpIndex Then
                
                'iterate categories
                For i3 = 1 To ArrSummary(i2).NumberOfCategories
                
                    '<==Alter Unique
                    'Slow Movers (any part that hasn't been seen and hasn't moved for 4 days from today)
                    If ArrWIP(i1).Notes.ValuesList(10) < (CDate(dToday - 4)) And ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCString = "" And ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCDate < (CDate(dToday - 4)) Then
                        'add UC to Category 1's UCList
                        ArrSummary(i2).CategoryList(1).AddUnitColumn ArrWIP(i1)
                        'skip for loops to next UC object
                        GoTo lineNextUC
                    'Slow Movers (hasn't moved for 4 days from today, but has been seen)
                    ElseIf ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCString = "" And ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCDate < (CDate(dToday - 4)) Then
                        'add UC to Category 1's UCList
                        ArrSummary(i2).CategoryList(1).AddUnitColumn ArrWIP(i1)
                        'skip for loops to next UC object
                        GoTo lineNextUC
                    '<==Alter Unique
                    
                    '<==Alter Unique
                    'Dark Greens, Oranges, Pinks counted as regular completes in specific
                    ElseIf ((i2 = 8) Or (i2 = 13) Or (i2 = 17) Or (i2 = 21) Or (i2 = 24) Or (i2 = 27)) And ((ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.OldOX_RCVD) Or (ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.NewOX_RCVD) Or (ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.SKOX_RCVD)) Then
                        'add UC to Category 3's UCList
                        ArrSummary(i2).CategoryList(3).AddUnitColumn ArrWIP(i1)
                        'skip for loops to next UC object
                        GoTo lineNextUC
                    '<==Alter Unique
                    
                    'Correct Summary Category found
                    ElseIf ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrSummary(i2).CategoryList(i3).Color Then
                        'add UC to Category's UCList
                        ArrSummary(i2).CategoryList(i3).AddUnitColumn ArrWIP(i1)
                        'skip for loops to next UC object
                        GoTo lineNextUC
                    
                    '<==Alter Unique
                    'All other colors counted as regular completes (excluding Dark-Color UCs)
                    ElseIf Not ((ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.OldOX_RCVD) Or (ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.NewOX_RCVD) Or (ArrWIP(i1).OperationsList(ArrWIP(i1).LastOpCompleted).UCColor = ArrWIP(i1).OperationsList(1).UCColorList.SKOX_RCVD)) Then
                        'add UC to Category 3's UCList
                        ArrSummary(i2).CategoryList(3).AddUnitColumn ArrWIP(i1)
                        'skip for loops to next UC object
                        GoTo lineNextUC
                    '<==Alter Unique
                    
                    End If
                
                Next i3
            
            End If
        
        Next i2
lineNextUC:
        'update loading bar
        lBar.UpdateLoadingBar "Organizing Summary Columns...", i1, UBound(ArrWIP)
    Next i1
    
    
    'clear summary
    Call ClearSummary(rTitleRow, rFirstListRow, cFirstSC, cLastSC, iCats)
    'write summary to sheet
    Call WriteSummary(rTitleRow, rFirstListRow, cFirstSC, cLastSC, iCats)
    

End Sub

Public Sub ClearSummary(ByVal rTitleRow As Integer, ByVal rFirstListRow As Integer, ByVal cFirstSC As Integer, ByVal cLastSC As Integer, ByVal iCats As Integer)

    Dim rngSummary As Range 'range object to clear entire summary
    Dim i As Integer 'iterator
    
    For i = 1 To 2
    
        'activate summary page
        If i = 1 Then: SheetSummary24K.Activate
        If i = 2 Then: SheetSummarySK.Activate
        
        'stop screen updates
        Application.ScreenUpdating = False
        
        Set rngSummary = Range(Cells((rTitleRow + 1), cFirstSC), Cells(2000, cLastSC))
        
        'clear values
        rngSummary.ClearContents
        'clear colors
        rngSummary.Interior.Color = xlNone
        'unbold text
        rngSummary.Font.Bold = False
        'make text black
        rngSummary.Font.Color = RGB(0, 0, 0)
        
        'start screen updates
        Application.ScreenUpdating = True
    
    Next i

End Sub

Public Sub WriteSummary(ByVal rTitleRow As Integer, ByVal rFirstListRow As Integer, ByVal cFirstSC As Integer, ByVal cLastSC As Integer, ByVal iCats As Integer)

    Dim tClrOpt As TC_CellColor_24K 'Color Options Object
    Dim i As Integer 'iterator
    Dim i1 As Integer 'iterator
    Dim i2 As Integer 'iterator
    Dim i3 As Integer 'iterator
    Dim iC As Integer 'counter for 24K summary entries
    Dim iCSK As Integer 'counter for SK summary entries
    
    
    'initialize color options object
    Set tClrOpt = New TC_CellColor_24K
    
    
    'activate summary page
    SheetSummary24K.Activate
    
    
    'stop screen updates
    Application.ScreenUpdating = False
    
    
    '<==Alter Unique
    'reset each category 1's color to red and SK colors
    For i1 = 1 To UBound(ArrSummary)
        ArrSummary(i1).CategoryList(1).Color = RGB(255, 0, 0)
        ArrSummary(i1).CategoryList(1).SKColor = RGB(255, 0, 0)
        ArrSummary(i1).CategoryList(2).SKColor = tClrOpt.RTO
        ArrSummary(i1).CategoryList(3).SKColor = tClrOpt.OldComplete
    Next i1
    '<==Alter Unique
    
    
    'iterate ArrSummary
    For i1 = 1 To UBound(ArrSummary)
    
        'reset iC counters
        iC = rFirstListRow
        iCSK = rFirstListRow
        
        'iterate categories
        For i2 = 1 To ArrSummary(i1).NumberOfCategories
            
            'write total
                'SK
                SheetSummarySK.Cells((rTitleRow + i2), ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).SKTotal
                If ArrSummary(i1).CategoryList(i2).SKTotal > 0 Then
                    SheetSummarySK.Cells((rTitleRow + i2), ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).SKColor
                End If
                '24K
                Cells((rTitleRow + i2), ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).Total
                If ArrSummary(i1).CategoryList(i2).Total > 0 Then
                    Cells((rTitleRow + i2), ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).Color
                End If
            
            'iterate UCList
            
                '24K
                For i3 = 1 To ArrSummary(i1).CategoryList(i2).Total
                
                    'write UC list Abbreviation
                    Cells(iC, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).UCList(i3).TNumAbbr
                    Cells(iC, ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).Color
                        '<==Alter Unique
                        If ArrSummary(i1).CategoryList(i2).UCList(i3).FIPL Then
                            Cells(iC, ((cFirstSC + i1) - 1)).Font.Color = RGB(255, 255, 0)
                            Cells(iC, ((cFirstSC + i1) - 1)).Font.Bold = True
                        End If
                        '<==Alter Unique
                    
                    'increment ic
                    iC = iC + 1
                    
                    'write UC list date
                    If ArrSummary(i1).CategoryList(i2).UCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).UCList(i3).LastOpCompleted).UCString = "" Then
                        Cells(iC, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).UCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).UCList(i3).LastOpCompleted).UCDate
                    'write UC list string
                    Else
                        Cells(iC, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).UCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).UCList(i3).LastOpCompleted).UCString
                    End If
                    'apply UC list entry color
                    Cells(iC, ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).Color
                        '<==Alter Unique
                        If ArrSummary(i1).CategoryList(i2).UCList(i3).FIPL Then
                            Cells(iC, ((cFirstSC + i1) - 1)).Font.Color = RGB(255, 255, 0)
                            Cells(iC, ((cFirstSC + i1) - 1)).Font.Bold = True
                        End If
                        '<==Alter Unique
                    
                    'increment ic again
                    iC = iC + 1
                
                Next i3
                
                'SK
                For i3 = 1 To ArrSummary(i1).CategoryList(i2).SKTotal
                
                    'write UC list Abbreviation
                    SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).SKUCList(i3).TNumAbbr
                    SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).SKColor
                        '<==Alter Unique
                        If ArrSummary(i1).CategoryList(i2).SKUCList(i3).FIPL Then
                            SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Font.Color = RGB(255, 255, 0)
                            SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Font.Bold = True
                        End If
                        '<==Alter Unique
                    
                    'increment ic sk
                    iCSK = iCSK + 1
                    
                    'write UC list date
                    If ArrSummary(i1).CategoryList(i2).SKUCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).SKUCList(i3).LastOpCompleted).UCString = "" Then
                        SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).SKUCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).SKUCList(i3).LastOpCompleted).UCDate
                    'write UC list string
                    Else
                        SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Value = ArrSummary(i1).CategoryList(i2).SKUCList(i3).OperationsList(ArrSummary(i1).CategoryList(i2).SKUCList(i3).LastOpCompleted).UCString
                    End If
                    'apply UC list entry color
                    SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Interior.Color = ArrSummary(i1).CategoryList(i2).SKColor
                        '<==Alter Unique
                        If ArrSummary(i1).CategoryList(i2).SKUCList(i3).FIPL Then
                            SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Font.Color = RGB(255, 255, 0)
                            SheetSummarySK.Cells(iCSK, ((cFirstSC + i1) - 1)).Font.Bold = True
                        End If
                        '<==Alter Unique
                    
                    'increment ic sk again
                    iCSK = iCSK + 1
                
                Next i3
            
        Next i2
        
        'update loading bar
        lBar.UpdateLoadingBar "Writing Summary Columns...", i1, UBound(ArrSummary)
    
    Next i1
    
    
    'insert black lines into WIP
    Call Summary_InsertBlackLines
    
    
    'start screen updates
    Application.ScreenUpdating = True


End Sub

Public Sub Summary_InsertBlackLines()

    Dim tUCol As TC_UnitColumn 'Temp UC object for reading where to place black lines
    Dim c As Range 'iterator
    Dim i As Integer 'iterator
    Dim cntBL As Integer 'counter
    
    'initialize counter
    cntBL = 0
    
    'iterate WIP array
    For i = 1 To UBound(ArrWIP)
        
        'grab UC object
        Set tUCol = ArrWIP(i)
        
        'grab c range
        Set c = SheetWIP.Range(tUCol.ColumnAddress)
        
        'adjust c range
        Set c = c.Offset(0, cntBL)
        
        'insert black lines
        If tUCol.Headers.ValuesList(4) = 36 Then
            'don't place black line if redline already there
            If Not c.Offset(0, 1).EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                'insert line after current line
                c.Offset(0, 1).EntireColumn.Insert
                'black out next line
                c.Offset(0, 1).EntireColumn.Interior.Color = RGB(0, 0, 0)
                'increment counter for accurate black lines placement
                cntBL = cntBL + 1
            End If
        End If
        
    Next i

End Sub

Public Sub WaterfallSort()

    Dim dummyCellColorObj As TC_CellColor_24K 'dummy object variable for grabbing array of color order indexes
    Dim arrOpRowGroup() As Variant 'Array for grabbing all similar last op UC objects for easier sorting
    Dim arrWtrfllWIP() As TC_UnitColumn 'Array for placing sorted waterfall WIP before assigning to arrWIP
    Dim i1 As Integer 'generic integer object for iteration
    Dim i2 As Integer 'generic integer object for iteration
    Dim i3 As Integer 'generic integer object for iteration
    Dim i4 As Integer 'generic integer object for iteration
    Dim o As Integer 'integer object for resizing arrOpRowGroup to the correct size
    Dim w As Integer 'integer object for assigning waterfall array order
    Dim wAdd As Integer 'integer object for adding current arrOpRowGroup size to wTotal
    Dim wTotal As Integer 'integer object for tracking current total size of waterfall array
    Dim setNum As Integer 'assigns Engine Set Number to each UC after Waterfalling
    Dim indWIPsection As Integer 'WIP section iterator (4 sections)
    
    'activate WIP worksheet
    SheetWIP.Activate
    
    'initialize variables
    Set dummyCellColorObj = New TC_CellColor_24K
    w = 0
    wAdd = 0
    wTotal = 0
    
    'iterate WIP sections
    For indWIPsection = 1 To 4
    
        'reset Engine Set Counter
        setNum = 0
        
        'iterate opRows
        For i1 = 1 To ArrWIP(1).NumberOfOps '<==Alter (if more than one part number i.e. SGVs)
            'initialize variables (to reset every op row looked at)
            o = 0
            'if op row is not hidden
            If ArrWIP(1).OperationsList(i1).Enabled Then
                
                
                'iterate arrwip for entries with this oprow as its last op completed
                For i2 = 1 To UBound(ArrWIP) '<==Alter (if more than one part number i.e. SGVs)
                    'UC's last op in oprow found (UC must be a part of the current WIP section)
                    If ArrWIP(i2).LastOpCompleted = i1 And (ArrWIP(i2).WIPSectionIndex = indWIPsection) Then '<==Alter (if more than one part number i.e. SGVs)
                        'increment o integer
                        o = o + 1
                        'redimension arrKeys
                        ReDim Preserve arrOpRowGroup(1 To 3, 1 To o)
                        'add UC object (array row 3) and UC object's lastopcompleted date (array row 2) and UC's ColorOrderIndex (array row 3)
                        Set arrOpRowGroup(3, o) = ArrWIP(i2) '<==Alter (if more than one part number i.e. SGVs)
                        If ArrWIP(i2).OperationsList(i1).UCString = "" Then
                            arrOpRowGroup(2, o) = ArrWIP(i2).OperationsList(i1).UCDate '<==Alter (if more than one part number i.e. SGVs)
                        ElseIf ArrWIP(i2).OperationsList(i1).UCString <> "" Then
                            arrOpRowGroup(2, o) = ArrWIP(i2).OperationsList(i1).UCString '<==Alter (if more than one part number i.e. SGVs)
                        End If
                        arrOpRowGroup(1, o) = ArrWIP(i2).ColorOrderIndex '<==Alter (if more than one part number i.e. SGVs)
                    End If
                Next i2
                
                
                'Only sort and apend to arrWtrfllWIP if the Op Row isn't empty
                If o <> 0 Then
                    
                    'Primary Sort: Color Order; Secondary Sort: Date Order
                    For i2 = UBound(dummyCellColorObj.arrColorOrder, 2) To 1 Step -1
                        'find entries with this color
                        For i3 = 1 To UBound(arrOpRowGroup, 2)
                            
                            
                            'entry with this color order index that has not been sorted
                            If (arrOpRowGroup(1, i3) = i2) And Not (arrOpRowGroup(3, i3).WSorted) Then
                                'reset .WSorted prop
                                arrOpRowGroup(3, i3).WSorted = True
                                'look for correct location (looking right to left)
                                For i4 = UBound(arrOpRowGroup, 2) To 1 Step -1
                                    
                                    'comparing entry to itself yields nothing
                                    If i4 = i3 Then
                                        'do nothing
                                        'i3 decrement not necessary here because no movement occurs
                                    
                                    'place after entry with same color order index and earlier date (or non-date)
                                    ElseIf (arrOpRowGroup(1, i4) = arrOpRowGroup(1, i3)) And ((arrOpRowGroup(2, i4) <= arrOpRowGroup(2, i3)) Or (Not IsDate(arrOpRowGroup(2, i3)))) Then
                                        'i3 < i4 => use i4 ; i3 > i4 => use (i4 + 1)
                                        If i3 < i4 Then
                                            arrOpRowGroup = MoveArrayEntry(arrOpRowGroup, i3, i4)
                                        ElseIf i3 > i4 Then
                                            arrOpRowGroup = MoveArrayEntry(arrOpRowGroup, i3, (i4 + 1))
                                        End If
                                        'decrement i3 back one in case of missed entries due to movement
                                        i3 = i3 - 1
                                        Exit For 'leave i4 for block
                                    
                                    'place after entry with lower color order index
                                    ElseIf arrOpRowGroup(1, i4) < arrOpRowGroup(1, i3) Then
                                        'i3 < i4 => use i4 ; i3 > i4 => use (i4 + 1)
                                        If i3 < i4 Then
                                            arrOpRowGroup = MoveArrayEntry(arrOpRowGroup, i3, i4)
                                        ElseIf i3 > i4 Then
                                            arrOpRowGroup = MoveArrayEntry(arrOpRowGroup, i3, (i4 + 1))
                                        End If
                                        'decrement i3 back one in case of missed entries due to movement
                                        i3 = i3 - 1
                                        Exit For 'leave i4 for block
                                    
                                    'no placement found before end of iteration, so place at front of array
                                    ElseIf i4 = 1 Then
                                        'always use i4 for moving to the first position
                                        arrOpRowGroup = MoveArrayEntry(arrOpRowGroup, i3, i4)
                                        'i3 decrement not necessary here because i3 will always be moving BACKWARDS to position 1
                                    
                                    End If
                                    
                                Next i4
                            End If
                            
                            
                        Next i3
                    Next i2
                    
                    'reset all UC objects' .WSorted property to False
                    For i2 = 1 To UBound(arrOpRowGroup, 2)
                        arrOpRowGroup(3, i2).WSorted = False
                    Next i2
                    
                    'grab size of current arrOpRowGroup
                    wAdd = UBound(arrOpRowGroup, 2)
                    
                    'add op row array to waterfall array
                    ReDim Preserve arrWtrfllWIP(1 To (wTotal + wAdd))
                    For i2 = 1 To wAdd
                        w = w + 1 'increment w to give UC's new waterfall index
                        'increment engine set number
                        If w <= 36 Then
                            setNum = w
                        ElseIf setNum <= 36 Then
                            setNum = setNum + 1
                        End If
                        'reset counter
                        If setNum = 37 Then
                            setNum = 1
                        End If
                        'add UC abbrv and object
                        Set arrWtrfllWIP((wTotal + i2)) = arrOpRowGroup(3, i2)
                        'set UC's waterfall index prop
                        arrWtrfllWIP((wTotal + i2)).WaterfallIndex = w
                        'set UC's Engine Set Number
                        arrWtrfllWIP((wTotal + i2)).Headers.ValuesList(4) = setNum
                    Next i2
                    
                    'increment total arrWtrfllWIP size
                    wTotal = wTotal + wAdd
                    
                    'reset op row array
                    Erase arrOpRowGroup
                
                End If
                
                
            End If
            
            'update loading bar
            lBar.UpdateLoadingBar "Waterfalling WIP section " & Str(indWIPsection) & " of 4...", i1, ArrWIP(1).NumberOfOps '<==Alter (if more than one part number i.e. SGVs)
            
        Next i1
    
    Next indWIPsection
    
    'assign waterfall array to WIP array
    ArrWIP = arrWtrfllWIP '<==Alter (if more than one part number i.e. SGVs)

End Sub

Public Function MoveArrayEntry(ByRef arrayReordering() As Variant, ByVal indexMoving As Integer, ByVal indexDestination As Integer) As Variant
'This function will move the given indexed entry to the indexed destination given within the array given.
    
    Dim t() As Variant 'copy of array given minus the moving index
    Dim M() As Variant 'only the moving index
    Dim arrFinal() As Variant 'finalized array ready to return from function
    Dim aMax As Integer 'number of rows in array
    Dim bMax As Integer 'number of columns in array
    Dim a As Integer 'integer object to iterate through all rows in one entry of the given array
    Dim b As Integer 'integer object to iterate through all columns in the given array
    Dim c As Integer 'integer to copy values into T
    
    
    'resize T to fit array given
    aMax = UBound(arrayReordering, 1)
    bMax = UBound(arrayReordering, 2)
    ReDim t(1 To aMax, 1 To (bMax - 1)) 'one less entry with indexMoving taken out
    ReDim M(1 To aMax) 'only the moving index
    ReDim arrFinal(1 To aMax, 1 To bMax) 'same size as entering array
    
    
    'copy array to T minus indexMoving
    c = 0 'initialize y
    For b = 1 To bMax 'b is index for giving array
        
        'not indexMoving
        If b <> indexMoving Then
            c = c + 1 'increment c
            For a = 1 To aMax
                'object variable
                If IsObject(arrayReordering(a, b)) Then
                    Set t(a, c) = arrayReordering(a, b)
                'regular variable
                Else
                    t(a, c) = arrayReordering(a, b)
                End If
            Next a
        
        'ignore indexMoving for T but put into M
        ElseIf b = indexMoving Then
            c = c 'do not increment c
            For a = 1 To aMax
                'object variable
                If IsObject(arrayReordering(a, b)) Then
                    Set M(a) = arrayReordering(a, b)
                'regular variable
                Else
                    M(a) = arrayReordering(a, b)
                End If
            Next a
        
        End If
    Next b
    
    
    'finalize array for return
    c = 0 'initialize y
    For b = 1 To bMax 'b is index for receiving array this time
        
        'not indexDestination
        If b <> indexDestination Then
            c = c + 1 'increment c
            For a = 1 To aMax
                'object variable
                If IsObject(t(a, c)) Then
                    Set arrFinal(a, b) = t(a, c)
                'regular variable
                Else
                    arrFinal(a, b) = t(a, c)
                End If
            Next a
        
        'copy M into destination index
        ElseIf b = indexDestination Then
            c = c 'do not increment c
            For a = 1 To aMax
                'object variable
                If IsObject(M(a)) Then
                    Set arrFinal(a, b) = M(a)
                'regular variable
                Else
                    arrFinal(a, b) = M(a)
                End If
            Next a
        
        End If
    Next b
    
    
    'return array
    MoveArrayEntry = arrFinal

End Function

Public Sub ClearWIP()

    Dim sWIP As Range 'wip section start
    Dim eWIP As Range 'wip section end
    Dim i As Integer 'iterator
    
    'activate WIP worksheet
    SheetWIP.Activate
    
    'stop screen updates
    Application.ScreenUpdating = False

    'clear WIP
    For i = 1 To 4
        'set start and end ranges for WIP section
        If i = 1 Then
            Set sWIP = StartWIP
            Set eWIP = EndWIP
        ElseIf i = 2 Then
            Set sWIP = StartWIP2
            Set eWIP = EndWIP2
        ElseIf i = 3 Then
            Set sWIP = StartWIP3
            Set eWIP = EndWIP3
        ElseIf i = 4 Then
            Set sWIP = StartWIP4
            Set eWIP = EndWIP4
        End If
        'clear WIP section
        SheetWIP.Range(sWIP, eWIP).Select
        Selection.ClearContents 'delete all cell values
        Selection.ClearComments 'delete all comments
        Selection.Interior.Color = xlNone 'white out all cells
    Next i
    
    'select top-left cell
    SheetWIP.Range("A1").Select
    
    'start screen updates
        'screen updates not started until after writing complete

End Sub

Public Sub WriteWIP()
'Waterfall Sort should have sorted all UC's based on their WIP section index.
'This means, that this sub should only have to worry about skipping redlines,
' all other orders should already have been taken into account.

    Dim tUCol As TC_UnitColumn 'Temporary Unit Column object for reading each array entry's data
    Dim tOpRow As TC_OpRow 'Temporary OpRow object for reading each array entry's OpRow data
    Dim c As Range 'generic range iteration object
    Dim i As Integer 'generic integer object for writing Unit Column objects from WIP array to WIP Tracker
    Dim j As Integer 'generic integer object (inside i blocks)
    Dim cntBlackLines As Integer 'counter for adding black lines and not writing UC info into them
    Dim cntRedLines As Integer 'counter for red lines and not writing UC info into them
    
    'activate WIP worksheet
    SheetWIP.Activate
    
    'initialize counter
    cntBlackLines = 0
    cntRedLines = 0
    
    'stop screen updates
        'screen updates are already off before ClearWIP is called
    
    'iterate array
    For i = 1 To UBound(ArrWIP) '<==Alter (if more than one part number i.e. SGVs)
    
    
        'grab current Unit Column object
        Set tUCol = ArrWIP(i) '<==Alter (if more than one part number i.e. SGVs)
        
        
        'Unit's Headers
        Set c = StartWIP.Offset(0, (tUCol.WaterfallIndex - 1 + cntBlackLines + cntRedLines)) 'set first cell in group
        For j = 1 To tUCol.Headers.GroupSize
            'values
            c.Offset((j - 1), 0).Value = tUCol.Headers.ValuesList(j)
            'colors
            c.Offset((j - 1), 0).Interior.Color = tUCol.Headers.ColorsList(j)
            'comments
            If Not tUCol.Headers.CommentsList(j) = "" Then
                c.Offset((j - 1), 0).AddComment
                c.Offset((j - 1), 0).Comment.Text tUCol.Headers.CommentsList(j)
            End If
        Next j
        
        
        'Unit's Operations Group
        Set c = c.Offset(tUCol.Headers.GroupSize, 0) 'set first cell in group
        For j = 1 To tUCol.NumberOfOps
            'set current op row object
            Set tOpRow = tUCol.OperationsList(j)
            'values
            If tOpRow.UCString <> "" Then 'add UC String if it isnt empty
                c.Offset((j - 1), 0).Value = tOpRow.UCString
            ElseIf tOpRow.UCDate <> CDate(0) Then 'only add value to cell if not the zero date
                c.Offset((j - 1), 0).Value = tOpRow.UCDate
            End If
            'colors
            c.Offset((j - 1), 0).Interior.Color = tOpRow.UCColor
            'comments
            If Not tOpRow.UCComment = "" Then
                c.Offset((j - 1), 0).AddComment
                c.Offset((j - 1), 0).Comment.Text tOpRow.UCComment
            End If
            'release op row object
            Set tOpRow = Nothing
        Next j
        
        
        'Unit's Notes Group
        Set c = c.Offset(tUCol.NumberOfOps, 0) 'set first cell in group
        For j = 1 To tUCol.Notes.GroupSize
            'values
            c.Offset((j - 1), 0).Value = tUCol.Notes.ValuesList(j)
            'colors
            c.Offset((j - 1), 0).Interior.Color = tUCol.Notes.ColorsList(j)
            'comments
            If Not tUCol.Notes.CommentsList(j) = "" Then
                c.Offset((j - 1), 0).AddComment
                c.Offset((j - 1), 0).Comment.Text tUCol.Notes.CommentsList(j)
            End If
        Next j
        
        
        'insert black line
        If tUCol.Headers.ValuesList(4) = 36 Then
            'don't place black line if redline already there
            If Not c.Offset(0, 1).EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                'insert line after current line
                c.Offset(0, 1).EntireColumn.Insert
                'black out next line
                c.Offset(0, 1).EntireColumn.Interior.Color = RGB(0, 0, 0)
                'increment black line counter to adjust WIP writing accordingly
                cntBlackLines = cntBlackLines + 1
            End If
        End If
        
        
        'check for redline in next column over
        If c.Offset(0, 1).EntireColumn.Interior.Color = RGB(255, 0, 0) Then
            'increment redline counter
            cntRedLines = cntRedLines + 1
        End If
        
        
        'release current Unit Column object
        Set tUCol = Nothing
        
        
        'update loading bar
        lBar.UpdateLoadingBar "Writing WIP to Tracker...", i, UBound(ArrWIP) '<==Alter (if more than one part number i.e. SGVs)
    
    
    Next i
    
    'start screen updates
    Application.ScreenUpdating = True
    
    'Delete Backup WIP Sheet
    Call DeleteWIPSheetCopy
    
    'Close Program
    End

End Sub

Public Sub CreateWIPSheetCopy()
'Creates a Copy of the WIP Sheet in case something goes wrong between ClearWIP and WriteWIP.

    'activate Original WIP Sheet
    SheetWIP.Activate
    
    'create copy of Original WIP Sheet
    SheetWIP.Copy Before:=SheetWIP
    
    'assign copy of Original WIP Sheet
    Set SheetWIP_BACKUP = ActiveSheet
    
    'rename Backup sheet
    SheetWIP_BACKUP.Name = SheetWIP.Name & " - BACKUP"
    
    'reactivate Original WIP Sheet
    SheetWIP.Activate

End Sub

Public Sub DeleteWIPSheetCopy()
'Deletes the Copy of the WIP Sheet after no errors occur between ClearWIP and WriteWIP.

    'activate Original WIP Sheet
    SheetWIP.Activate

    'delete backup copy without alearts
    Application.DisplayAlerts = False
    SheetWIP_BACKUP.Delete
    Application.DisplayAlerts = True

End Sub
