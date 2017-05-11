VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FOPS_SNRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pListIndex As Integer 'Index of the SN Row within the List Table
Private pWaterfallIndex As Integer 'Index count of green completed process steps for each SN Row
Private pDateFromPWA As Date 'Date Received from PWA. Secondary waterfall sort.
Private pWSorted As Boolean 'True: SN Row has been waterfalled already ; False: still needs to be waterfalled
Private pGroupSize As Integer 'Number of columns that exist in this group
Private pNumberOfOperations As Integer 'Number of Operation columns in the table
Private pColOpsStart As Integer 'Column number where operation columns start
Private pColOpsEnd As Integer 'Column number where operation columns end
Private pValuesList() As Variant 'Array of all values in group
Private pColorsList() As Variant 'Array of all cell colors in group
Private pCommentsList() As Variant 'Array of all cell comments in group

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' List Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ListIndex() As Integer
    ListIndex = pListIndex
End Property
Public Property Let ListIndex(Value As Integer)
    pListIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Waterfall Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get WaterfallIndex() As Integer
    WaterfallIndex = pWaterfallIndex
End Property
Public Property Let WaterfallIndex(Value As Integer)
    pWaterfallIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Date Received from PWA
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get DateFromPWA() As Date
    DateFromPWA = pDateFromPWA
End Property
Public Property Let DateFromPWA(Value As Date)
    pDateFromPWA = Value
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
' Group Size
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get GroupSize() As Integer
    GroupSize = pGroupSize
End Property
Public Property Let GroupSize(Value As Integer)
    pGroupSize = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of Operations
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NumberOfOperations() As Integer
    NumberOfOperations = pNumberOfOperations
End Property
Public Property Let NumberOfOperations(Value As Integer)
    pNumberOfOperations = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Column where Operations Begin
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColOpsStart() As Integer
    ColOpsStart = pColOpsStart
End Property
Public Property Let ColOpsStart(Value As Integer)
    pColOpsStart = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Column where Operations End
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColOpsEnd() As Integer
    ColOpsEnd = pColOpsEnd
End Property
Public Property Let ColOpsEnd(Value As Integer)
    pColOpsEnd = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Values Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ValuesList(ByVal indexVal As Integer) As Variant
    ValuesList = pValuesList(indexVal)
End Property
Public Property Let ValuesList(ByVal indexVal As Integer, Value As Variant)
    pValuesList(indexVal) = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Colors Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorsList(ByVal indexClr As Integer) As Variant
    ColorsList = pColorsList(indexClr)
End Property
Public Property Let ColorsList(ByVal indexClr As Integer, Value As Variant)
    pColorsList(indexClr) = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Comments Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CommentsList(ByVal indexCom As Integer) As Variant
    CommentsList = pCommentsList(indexCom)
End Property
Public Property Let CommentsList(ByVal indexCom As Integer, Value As Variant)
    pCommentsList(indexCom) = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Public Sub GrabData(ByVal StartingCell As Range, ByVal NumberOfColumns As Integer)

    Dim c As Range 'iterator
    Dim i As Integer 'index for adding values to arrays
    Dim clrWhite As Long 'White color long
    Dim clrBlack As Long 'Black color long
    Dim bWIndexFound As Boolean 'True: Waterfall index found ; False: Waterfall index not found yet
    
    'initialize colors
    clrWhite = RGB(255, 255, 255)
    clrBlack = RGB(0, 0, 0)
    
    'initialize booleans
    pWSorted = False
    
    'initialize waterfall index and boolean
    pWaterfallIndex = 0
    bWIndexFound = False
    
    'assign group size
    pGroupSize = NumberOfColumns
    
    'resize arrays
    ReDim pValuesList(1 To NumberOfColumns) As Variant
    ReDim pColorsList(1 To NumberOfColumns) As Variant
    ReDim pCommentsList(1 To NumberOfColumns) As Variant
    
    'initialize indexer
    i = 0
    
    'iterate group to grab data
    For Each c In Range(StartingCell, StartingCell.Offset(0, (NumberOfColumns - 1)))
    
        'increment indexer
        i = i + 1
        
        'grab values
        If Left(c.Formula, 1) = "=" Then
            pValuesList(i) = c.Formula
        Else
            pValuesList(i) = c.Value
        End If
        
        'grab colors
        pColorsList(i) = c.Interior.Color
        
        'set waterfall index (sort based on first white cell location)
            If Not bWIndexFound And c.Column >= pColOpsStart And c.Column <= pColOpsEnd And c.Interior.Color = clrWhite Then
                'set index
                pWaterfallIndex = (c.Column - pColOpsStart) + 1
                'index found, change boolean
                bWIndexFound = True
            'Closed part
            ElseIf Not bWIndexFound And c.Column > pColOpsEnd And c.Interior.Color = clrBlack Then
                'set index
                pWaterfallIndex = (c.Column - pColOpsStart) + 1
                'index found, change boolean
                bWIndexFound = True
            End If
        
        'grab comments
        If Not c.Comment Is Nothing Then
            pCommentsList(i) = c.Comment.Text
        Else
            pCommentsList(i) = "" 'no comment means add an empty string
        End If
    
    Next c

End Sub

