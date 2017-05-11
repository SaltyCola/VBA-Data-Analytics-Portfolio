VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_UnitColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pWIPSectionIndex As Integer 'Unit Column's WIP Section index number (1:WIP, 2:RTO, 3:New, 4:SK)

Private pColumnNumber As Double 'Unit Column's column number
Private pColumnAddress As String 'Unit Column's Tracking Number address
Private pColorOrderIndex As Integer 'Indexes Unit Column's last op completed color for waterfall order
Private pWSorted As Boolean 'tracks whether a UC has been sorted already in the waterfalling subroutine
Private pWaterfallIndex As Double 'Indexes all Unit Columns for proper display in waterfall order
Private pNumberOfOps As Integer 'The number of operations in this Unit's manufacturing process

Private pPartNumber As String 'Unit's Part Type Number
Private pTrackingNumber As String 'Unit's primary tracking number (24K = Serial Number)
Private pTNumAbbr As String 'The last 5 characters of the unit's primary tracking number
Private pLastOpCompleted As Integer 'Holds the most recently completed opRow's Index value
Private pLastDateSeen As Date 'Holds the date of the last time this UC was seen.
Private pFIPL As Boolean 'True: This Unit Column is currently on the Final Inspection Priority List (the UC's SN cell is colored Yellow) ; False: It is not on the FIPL

Private pHeaders As TC_UnitHeaderFooter 'Object containing all Unit's Header information
Private pOperationsList() As TC_OpRow 'Array containing all Unit's Operation Row objects
Private pNotes As TC_UnitHeaderFooter 'Object containing all Unit's Notes information below the operations list

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' WIP Section Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get WIPSectionIndex() As Integer
    WIPSectionIndex = pWIPSectionIndex
End Property
Public Property Let WIPSectionIndex(Value As Integer)
    pWIPSectionIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column Number
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColumnNumber() As Double
    ColumnNumber = pColumnNumber
End Property
Public Property Let ColumnNumber(Value As Double)
    pColumnNumber = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column Address
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColumnAddress() As String
    ColumnAddress = pColumnAddress
End Property
Public Property Let ColumnAddress(Value As String)
    pColumnAddress = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Color Order Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorOrderIndex() As Integer
    ColorOrderIndex = pColorOrderIndex
End Property
Public Property Let ColorOrderIndex(Value As Integer)
    pColorOrderIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Waterfall Sorted
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get WSorted() As Boolean
    WSorted = pWSorted
End Property
Public Property Let WSorted(Value As Boolean)
    pWSorted = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Waterfall Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get WaterfallIndex() As Double
    WaterfallIndex = pWaterfallIndex
End Property
Public Property Let WaterfallIndex(Value As Double)
    pWaterfallIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of Unit's Operations
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NumberOfOps() As Integer
    NumberOfOps = pNumberOfOps
End Property
Public Property Let NumberOfOps(Value As Integer)
    pNumberOfOps = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Part Number
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get PartNumber() As String
    PartNumber = pPartNumber
End Property
Public Property Let PartNumber(Value As String)
    pPartNumber = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Tracking Number
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TrackingNumber() As String
    TrackingNumber = pTrackingNumber
End Property
Public Property Let TrackingNumber(Value As String)
    pTrackingNumber = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Abbreviated Tracking Number
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TNumAbbr() As String
    TNumAbbr = pTNumAbbr
End Property
Public Property Let TNumAbbr(Value As String)
    pTNumAbbr = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit's Last Completed Operation Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LastOpCompleted() As Integer
    LastOpCompleted = pLastOpCompleted
End Property
Public Property Let LastOpCompleted(Value As Integer)
    pLastOpCompleted = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit's Last Date Seen
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LastDateSeen() As Date
    LastDateSeen = pLastDateSeen
End Property
Public Property Let LastDateSeen(Value As Date)
    pLastDateSeen = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit is on the Final Inspection Priority List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FIPL() As Boolean
    FIPL = pFIPL
End Property
Public Property Let FIPL(Value As Boolean)
    pFIPL = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Headers
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Headers() As TC_UnitHeaderFooter
    Set Headers = pHeaders
End Property
Public Property Let Headers(Value As TC_UnitHeaderFooter)
    Set pHeaders = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operations List Collection
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OperationsList(ByVal indexOpRow As Integer) As TC_OpRow
    Set OperationsList = pOperationsList(indexOpRow)
End Property
Public Property Let OperationsList(ByVal indexOpRow As Integer, Value As TC_OpRow)
    Set pOperationsList(indexOpRow) = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Notes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Notes() As TC_UnitHeaderFooter
    Set Notes = pNotes
End Property
Public Property Let Notes(Value As TC_UnitHeaderFooter)
    Set pNotes = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Private Sub Class_Initialize()

    'initialize Child Class Objects
    Set pHeaders = New TC_UnitHeaderFooter
    pHeaders.TitleColumn = 4 '<==Alter
    Set pNotes = New TC_UnitHeaderFooter
    pNotes.TitleColumn = 4 '<==Alter

End Sub

Public Sub GrabOperationsData(ByVal FinalOpCell As Range, ByVal NumberOfRows As Integer)
    
    Dim c As Range 'iterator
    Dim i As Integer 'generic index
    Dim tOpRow As TC_OpRow 'temporary Op Row object for storage in Operations List collection
    Dim bVisible As Boolean 'boolean to store whether row is visible or not
    
    'assign properties
    pNumberOfOps = NumberOfRows
    pLastOpCompleted = NumberOfRows
    
    'resize op row list array
    ReDim pOperationsList(1 To NumberOfRows) As TC_OpRow
    
    'iterate through range in reverse in order to easily grab last op completed
    For i = NumberOfRows To 1 Step -1
    
        'assign range object
        Set c = FinalOpCell.Offset((i - 1), 0)
        
        'test for row visibility
        If c.EntireRow.Hidden Then
            bVisible = False
        Else
            bVisible = True
        End If
        
        'create opRow object
        Set tOpRow = New TC_OpRow
        
        'grab info
        tOpRow.GrabRowData c.Row, 4, i, bVisible, c.Column '<==Alter
        
        'Grab last completed op index (non-hidden row)
        If tOpRow.UCColor <> tOpRow.UCColorList.Blank And tOpRow.UCColor <> tOpRow.UCColorList.FinalSST And bVisible Then
            pLastOpCompleted = i
        End If
        
        'store into array
        Set pOperationsList(i) = tOpRow
    
    Next i

End Sub
