VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pColumnNumber As Double 'Unit Column's column number
Private pColumnAddress As String 'Unit Column's column address
Private pColorOrderIndex As Integer 'Indexes Unit Column's last op completed color for waterfall order
Private pWSorted As Boolean 'tracks whether a UC has been sorted already in the waterfalling subroutine
Private pWaterfallIndex As Double 'Indexes all Unit Columns for proper display in waterfall order
Private pNumberOfOps As Integer 'The number of operations in this Unit's manufacturing process

Private pPartNumber As String 'Unit's Part Type Number
Private pTrackingNumber As String 'Unit's primary tracking number (30K = Serial Number)
Private pTNumAbbr As String 'The last 5 characters of the unit's primary tracking number
Private pLastOpCompleted As Integer 'Holds the most recently completed opRow's Index value

Private pHeaders As UnitHeaderFooter 'Object containing all Unit's Header information
Private pOperationsList As Collection 'Collection containing all Unit's Operation Row objects
Private pNotes As UnitHeaderFooter 'Object containing all Unit's Notes information below the operations list

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
' Headers
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Headers() As UnitHeaderFooter
    Set Headers = pHeaders
End Property
Public Property Let Headers(Value As UnitHeaderFooter)
    Set pHeaders = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operations List Collection
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OperationsList() As Collection
    Set OperationsList = pOperationsList
End Property
Public Property Let OperationsList(Value As Collection)
    Set pOperationsList = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Notes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Notes() As UnitHeaderFooter
    Set Notes = pNotes
End Property
Public Property Let Notes(Value As UnitHeaderFooter)
    Set pNotes = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Private Sub Class_Initialize()

    'initialize Child Class Objects
    Set pHeaders = New UnitHeaderFooter
    Set pOperationsList = New Collection
    Set pNotes = New UnitHeaderFooter

End Sub

Public Sub GrabOperationsData(ByVal FinalOpCell As Range, ByVal NumberOfRows As Integer)
    
    Dim c As Range 'generic iteration range object
    Dim i As Integer 'generic index
    Dim tOpRow As OpRow 'temporary Op Row object for storage in Operations List collection
    Dim bVisible As Boolean 'boolean to store whether row is visible or not
    
    'assign properties
    pNumberOfOps = NumberOfRows
    
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
        Set tOpRow = New OpRow
        
        'grab info
        tOpRow.GrabRowData c.Row, 2, i, bVisible, c.Column '<==Altered
        
        'Grab last completed op index
        If tOpRow.UCColor <> tOpRow.UCColorList.Blank Then
            pLastOpCompleted = i
        End If
        
        'store into collection
        If pOperationsList.Count = 0 Then 'store first entry
            pOperationsList.Add tOpRow, tOpRow.Title
        Else 'store remaining entries always at the front of the collection
            pOperationsList.Add tOpRow, tOpRow.Title, 1
        End If
    
    Next i

End Sub
