VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitHeaderFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pGroupSize As Integer 'Number of rows that exist in this group
Private pValuesColl As Collection 'Collection of all values in group
Private pColorsColl As Collection 'Collection of all cell colors in group
Private pCommentsColl As Collection 'Collection of all cell comments in group

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
' Values Collection
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ValuesColl() As Collection
    Set ValuesColl = pValuesColl
End Property
Public Property Let ValuesColl(Value As Collection)
    Set pValuesColl = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Colors Collection
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ColorsColl() As Collection
    Set ColorsColl = pColorsColl
End Property
Public Property Let ColorsColl(Value As Collection)
    Set pColorsColl = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Comments Collection
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CommentsColl() As Collection
    Set CommentsColl = pCommentsColl
End Property
Public Property Let CommentsColl(Value As Collection)
    Set pCommentsColl = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Private Sub Class_Initialize()

    'initialize Child Class Objects
    Set pValuesColl = New Collection
    Set pColorsColl = New Collection
    Set pCommentsColl = New Collection

End Sub

Public Sub GrabData(ByVal StartingCell As Range, ByVal NumberOfRows As Integer)

    Dim c As Range 'generic range iteration object
    Dim i As Integer 'index for adding values to collections
    
    'assign group size
    pGroupSize = NumberOfRows
    
    'initialize indexer
    i = 0
    
    'iterate group to grab data
    For Each c In Range(StartingCell, StartingCell.Offset(NumberOfRows - 1))
    
        'increment indexer
        i = i + 1
    
        'grab values
        If Left(c.Formula, 1) = "=" Then
            pValuesColl.Add c.Formula, Str(i)
        Else
            pValuesColl.Add c.Value, Str(i)
        End If
        
        'grab colors
        pColorsColl.Add c.Interior.Color, Str(i)
        
        'grab comments
        If Not c.Comment Is Nothing Then
            pCommentsColl.Add c.Comment.Text, Str(i)
        Else
            pCommentsColl.Add "", Str(i) 'no comment means add an empty string
        End If
    
    Next c

End Sub
