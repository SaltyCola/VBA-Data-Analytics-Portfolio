VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_SummaryCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pTitle As String 'The summary category's title
Private pTotal As Integer 'The number of Unit Columns that belong to this category
Private pColor As Long 'The color of UC last op completed this category will contain
Private pUCList() As TC_UnitColumn 'Array of Unit Column objects that belong to this Summary Category

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Summary Category's Title
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Title() As String
    Title = pTitle
End Property
Public Property Let Title(Value As String)
    pTitle = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Total number of Unit Columns
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Total() As Integer
    Total = pTotal
End Property
Public Property Let Total(Value As Integer)
    pTotal = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Color of this Category
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Color() As Long
    Color = pColor
End Property
Public Property Let Color(Value As Long)
    pColor = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Array: Unit Column List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCList(ByVal indexUC As Integer) As TC_UnitColumn
    Set UCList = pUCList(indexUC)
End Property
Public Property Let UCList(ByVal indexUC As Integer, Value As TC_UnitColumn)
    Set pUCList(indexUC) = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Public Sub AddUnitColumn(ByVal tUnitColumn As TC_UnitColumn)

    'increment category total
    pTotal = pTotal + 1
    
    'redim array
    ReDim Preserve pUCList(1 To pTotal) As TC_UnitColumn
    
    'add UC to array
    Set pUCList(pTotal) = tUnitColumn

End Sub
