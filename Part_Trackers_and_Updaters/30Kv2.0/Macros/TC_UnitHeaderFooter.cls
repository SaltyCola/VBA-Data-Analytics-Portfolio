VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_UnitHeaderFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pGroupSize As Integer 'Number of rows that exist in this group
Private pTitleColumn As Integer 'Column Number where the row titles are
Private pTitlesList() As String 'Array of all row titles in group
Private pValuesList() As Variant 'Array of all values in group
Private pColorsList() As Variant 'Array of all cell colors in group
Private pCommentsList() As Variant 'Array of all cell comments in group

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
' Title Column
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TitleColumn() As Integer
    TitleColumn = pTitleColumn
End Property
Public Property Let TitleColumn(Value As Integer)
    pTitleColumn = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Titles Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TitlesList(ByVal indexTitle As Integer) As String
    TitlesList = pTitlesList(indexTitle)
End Property
Public Property Let TitlesList(ByVal indexTitle As Integer, Value As String)
    pTitlesList(indexTitle) = Value
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

Public Sub GrabData(ByVal StartingCell As Range, ByVal NumberOfRows As Integer)

    Dim c As Range 'generic range iteration object
    Dim i As Integer 'index for adding values to arrays
    
    'assign group size
    pGroupSize = NumberOfRows
    
    'resize arrays
    ReDim pTitlesList(1 To NumberOfRows) As String
    ReDim pValuesList(1 To NumberOfRows) As Variant
    ReDim pColorsList(1 To NumberOfRows) As Variant
    ReDim pCommentsList(1 To NumberOfRows) As Variant
    
    'initialize indexer
    i = 0
    
    'iterate group to grab data
    For Each c In Range(StartingCell, StartingCell.Offset(NumberOfRows - 1))
    
        'increment indexer
        i = i + 1
        
        'grab titles
        pTitlesList(i) = Cells(c.Row, pTitleColumn).Value
    
        'grab values
        If Left(c.Formula, 1) = "=" Then
            pValuesList(i) = c.Formula
        Else
            pValuesList(i) = c.Value
        End If
        
        'grab colors
        pColorsList(i) = c.Interior.Color
        
        'grab comments
        If Not c.Comment Is Nothing Then
            pCommentsList(i) = c.Comment.Text
        Else
            pCommentsList(i) = "" 'no comment means add an empty string
        End If
    
    Next c

End Sub
