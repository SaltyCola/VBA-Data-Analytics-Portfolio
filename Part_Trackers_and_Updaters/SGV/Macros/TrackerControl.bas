Attribute VB_Name = "TrackerControl"
'To Alter the code to work for other part numbers,
'and different tracker organizations search for: <==Alter
'Once completed, change <==Alter to <==Altered

'Also, must create new CellColor Class for each new part number


'If a color is added then the CellColor Class must be updated accordingly.
'Any rows added must be updated within this module.


'Declarations
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

'Public Variables
Public BookWIP As Workbook 'Workbook containing WIP
Public SheetWIP As Worksheet 'Worksheet containing WIP
Public ArrWIP() As Variant 'Array to store all Unit Columns in WIP '<==Altered (if more than one part number, i.e. SGVs, must erase ArrWIP at the start of each new part number before ReadWIP)
Public Arr080WIP() As Variant 'Array to store all Unit Columns in 080 WIP '<==Altered (unique)
Public Arr180WIP() As Variant 'Array to store all Unit Columns in 180 WIP '<==Altered (unique)
Public Arr280WIP() As Variant 'Array to store all Unit Columns in 280 WIP '<==Altered (unique)
Public Arr380WIP() As Variant 'Array to store all Unit Columns in 380 WIP '<==Altered (unique)
Public Arr480WIP() As Variant 'Array to store all Unit Columns in 480 WIP '<==Altered (unique)
Public StartWIP As Range 'first cell of WIP Range (top left corner)
Public EndWIP As Range 'last cell of WIP Range (bottom right corner)
Public LBar As LoadingBar 'Loading Bar that will be used throughout program
Public LInt As Integer 'Loading Bar integer that will keep track of which img should be visible

Public Sub TrackerUpdaterProgram()

    Dim cRedLine As Range 'range iterator for finding EndWIP range
    Dim rowsWIP As Integer 'integer that must be manually set to hold the final row in the WIP grid
    Dim b1stVisible As Boolean 'boolean for grabbing StartWIP position based on first unhidden Unit Column
    
    Dim iSGV As Integer 'integer iterator for SGV part number WIP tabs '<==Altered (unique)
    
    'iterate SGV tabs '<==Altered (unique)
    For iSGV = 0 To 4 '<==Altered (unique)
    
        'initialize public variables
        Set BookWIP = Workbooks("PWAA SGV WIP Status and Detail Tracking") '<==Altered
        Set SheetWIP = BookWIP.Worksheets("5319" & iSGV & "80") '<==Altered
        
        'initialize private variables
        rowsWIP = 40 '<==Altered
        b1stVisible = True
        
        'initialize loading bar
        Set LBar = New LoadingBar
        For LInt = 1 To 166
            LBar.Controls("Image" & LInt).Visible = False
        Next LInt
        LInt = 0
        LBar.Controls("Image" & LInt).Visible = True
        
        'stop screen updates
        Application.ScreenUpdating = False
        
        'activate WIP worksheet
        SheetWIP.Activate
    
        'initialize StartWIP and EndWIP variables
        For Each cRedLine In Range("1:1")
            'first visible unit column found
            If (b1stVisible) And (cRedLine.Column > 2) And Not (cRedLine.EntireColumn.Hidden) Then '<==Altered
                Set StartWIP = cRedLine
                'change boolean to stop looking for StartWIP position
                b1stVisible = False
            End If
            'redline found
            If cRedLine.EntireColumn.Interior.Color = RGB(255, 0, 0) Then
                'assign end of WIP
                Set EndWIP = Cells(rowsWIP, (cRedLine.Column - 1))
                Exit For 'end for loop
            End If
        Next cRedLine
        
        'call WIP subs
        Call ReadWIP 'requires loading bar
        
        'Assign general ArrWIP to corresponding part number ArrWIP '<==Altered (unique)
        If iSGV = 0 Then: Arr080WIP = ArrWIP '<==Altered (unique)
        If iSGV = 1 Then: Arr180WIP = ArrWIP '<==Altered (unique)
        If iSGV = 2 Then: Arr280WIP = ArrWIP '<==Altered (unique)
        If iSGV = 3 Then: Arr380WIP = ArrWIP '<==Altered (unique)
        If iSGV = 4 Then: Arr480WIP = ArrWIP '<==Altered (unique)
        
        'reset ArrWIP for next part number '<==Altered (unique)
        Erase ArrWIP() '<==Altered (unique)
        
    Next iSGV '<==Altered (unique)

End Sub

Public Sub UpdateLoadingBar(ByVal LoadingMessage As String, ByVal indexCurrent As Integer, ByVal indexTotal As Integer)

    Dim i As Integer 'generic integer variable for iteration
    
    'update message
    LBar.Label1.Caption = LoadingMessage
    
    'calculate new LInt
    LInt = Int((indexCurrent / indexTotal) * 166)
    
    'update loading image
    For i = 0 To 166
        LBar.Controls("Image" & i).Visible = False
    Next i
    LBar.Controls("Image" & LInt).Visible = True
    
    'show updates and maintain progression of code
    LBar.Show vbModeless
    DoEvents
    
    'close loading bar if end reached
    If indexCurrent = indexTotal Then
        LBar.Hide
    End If

End Sub

Public Sub ReadWIP()

    Dim arrTemp() As Variant 'Temporary Array to store all Unit Columns in WIP
    Dim tUCol As UnitColumn 'Temporary Unit Column object for entering data into array
    Dim tOpRow As OpRow 'Temporary Op Row object to add to unit column's operations list
    Dim c As Range 'generic range iteration object
    Dim i As Integer 'generic integer object for adding Unit Column objects to WIP array
    
    'initialize variables
    Set tUCol = New UnitColumn
    i = 0
    
    'iterate WIP
    For Each c In Range(Cells(13, StartWIP.Column), Cells(13, EndWIP.Column)) '<==Altered
    
        'only WIP if not hidden
        If Not c.EntireColumn.Hidden Then
            
            'redimension array
            i = i + 1
            ReDim Preserve arrTemp(1 To 2, 1 To i) As Variant
            
            'grab Unit Column property values
            tUCol.ColumnAddress = c.Address
            tUCol.ColumnNumber = c.Column
            tUCol.PartNumber = SheetWIP.Name '<==Altered
            tUCol.TrackingNumber = c.Value
            tUCol.TNumAbbr = Right(c.Value, 5)
            tUCol.WSorted = False
            'initialize indexes
            tUCol.ColorOrderIndex = 0
            tUCol.WaterfallIndex = i
            'GrabData Methods
            tUCol.Headers.GrabData Cells(1, c.Column), 19 '<==Altered
            tUCol.GrabOperationsData Cells(20, c.Column), 16 '<==Altered
            tUCol.Notes.GrabData Cells(36, c.Column), 5 '<==Altered
            'grab color order index
            tUCol.ColorOrderIndex = tUCol.OperationsList.Item(tUCol.LastOpCompleted).UCColorOrderIndex
            
            'add Unit to array
            arrTemp(1, i) = tUCol.TNumAbbr
            Set arrTemp(2, i) = tUCol
            
            'reset tUCol object
            Set tUCol = Nothing
            Set tUCol = New UnitColumn
            
            'update loading bar
            Call UpdateLoadingBar("Reading WIP from " & Right(SheetWIP.Name, 3) & " Tracker...", (c.Column - StartWIP.Column), (EndWIP.Column - StartWIP.Column))
            
        End If
    
    Next c
    
    'assign temp array to Public WIP array
    ArrWIP = arrTemp '<==Altered (if more than one part number i.e. SGVs)

End Sub

Public Sub ClearSGVSummary()
    
    'turn off screen updating
    Application.ScreenUpdating = False
    
    'activate WIP Summary Tab
    Worksheets("WIP Summary").Activate
    
    'clear all summary cells
    Worksheets("WIP Summary").Range("B5:R409").ClearContents
    Worksheets("WIP Summary").Range("B5:R409").Interior.Color = xlNone
    
    'turn on screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub UpdateSGVSummary()
'Double Click Event Code is located in the "Sheet12 (WIP Summary)" Code

    Dim i1 As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim c As Range
    Dim tSC As SummaryColumn_SGV 'temp SC object for creating and storing into Summary Array
    Dim SummaryArray(1 To 17, 1 To 5) As SummaryColumn_SGV 'array storing grid of SC objects to paste into Summary Page
    Dim tWIParr() As Variant 'temp array to hold current WIP array
    Dim tUC As UnitColumn 'temp UC object used to pull info from WIP array
    Dim strtCell As Range 'starting place of first entry of each new SC's UCList
    
    'clear summary
    Call ClearSGVSummary
        
    'stop screen updates
    Application.ScreenUpdating = False
    
    'Read part number WIPs
    Call TrackerUpdaterProgram
    
    'activate summary sheet
    Worksheets("WIP Summary").Activate
        
    'start screen updates
    Application.ScreenUpdating = True
    
    
    'initialize SummaryArray
    For i1 = 1 To 17
        For i2 = 1 To 5
            'create new object
            Set tSC = New SummaryColumn_SGV
            'set props
            tSC.SGVPartNumIndex = i2
            tSC.OpColumnIndex = i1
            tSC.InitializeSC
            'add to Summary Array
            Set SummaryArray(i1, i2) = tSC
        Next i2
    Next i1
    
    
    'iterate all SC Objects
    For i1 = 1 To 17
        For i2 = 1 To 5
        
            'grab SC
            Set tSC = SummaryArray(i1, i2)
            
            'get WIP array
            If i2 = 1 Then: tWIParr = Arr080WIP
            If i2 = 2 Then: tWIParr = Arr180WIP
            If i2 = 3 Then: tWIParr = Arr280WIP
            If i2 = 4 Then: tWIParr = Arr380WIP
            If i2 = 5 Then: tWIParr = Arr480WIP
            
            'iterate wip array
            For i3 = 1 To UBound(tWIParr, 2)
            
                'grab UC
                Set tUC = tWIParr(2, i3)
                
                'test for membership to SC
                If (tUC.ColorOrderIndex = tSC.OpColorOrderIndex) And (tUC.LastOpCompleted = tSC.OpIndex) Then
                    tSC.AddListEntry tUC.TNumAbbr, tUC.OperationsList(tUC.LastOpCompleted).UCDate
                ElseIf (tUC.ColorOrderIndex = 2) And (tSC.OpIndex = 0) Then 'repair cell SC
                    tSC.AddListEntry tUC.TNumAbbr, tUC.OperationsList(tUC.LastOpCompleted).UCDate
                End If
            
            Next i3
        
        Next i2
    Next i1
    
    
    'paste to Summary
    For i1 = 1 To 17
        For i2 = 1 To 5
        
            'grab SC
            Set tSC = SummaryArray(i1, i2)
            
            'paste total if not zero
            If tSC.Total <> 0 Then
                ActiveSheet.Cells((tSC.SGVPartNumIndex + 4), (tSC.OpColumnIndex + 1)).Value = tSC.Total
                ActiveSheet.Cells((tSC.SGVPartNumIndex + 4), (tSC.OpColumnIndex + 1)).Interior.Color = tSC.SGVColor
            End If
            
            'find first empty cell
            ActiveSheet.Columns(tSC.OpColumnIndex + 1).EntireColumn.Select
            For Each c In Selection
                If (c.Row > 9) And (IsEmpty(c)) Then
                    Set strtCell = c
                    Exit For
                End If
            Next c
            
            'paste list
            For i3 = 1 To tSC.Total
                strtCell.Offset(((i3 * 2) - 2), 0).Value = tSC.UCList(1, i3)
                strtCell.Offset(((i3 * 2) - 2), 0).Interior.Color = tSC.SGVColor
                strtCell.Offset(((i3 * 2) - 1), 0).Value = tSC.UCList(2, i3)
                strtCell.Offset(((i3 * 2) - 1), 0).Interior.Color = tSC.SGVColor
            Next i3
        
        Next i2
    Next i1
    

End Sub
