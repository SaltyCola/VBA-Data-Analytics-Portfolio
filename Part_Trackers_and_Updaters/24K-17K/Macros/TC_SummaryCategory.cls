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
Private pSKTotal As Integer 'The number of Unit Columns that belong to this category on the SK Summary Page
Private pColor As Long 'The color of UC last op completed this category will contain
Private pSKColor As Long 'The color of UC last op completed this category will contain on the SK Summary Page
Private pUCList() As TC_UnitColumn 'Array of Unit Column objects that belong to this Summary Category
Private pSKUCList() As TC_UnitColumn 'Array of Unit Column objects that belong to this Summary Category on the SK Summary Page

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
' Total number of Unit Columns (SK)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SKTotal() As Integer
    SKTotal = pSKTotal
End Property
Public Property Let SKTotal(Value As Integer)
    pSKTotal = Value
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
' Color of this Category (SK)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SKColor() As Long
    SKColor = pSKColor
End Property
Public Property Let SKColor(Value As Long)
    pSKColor = Value
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Array: Unit Column List (SK)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SKUCList(ByVal indexUC As Integer) As TC_UnitColumn
    Set SKUCList = pSKUCList(indexUC)
End Property
Public Property Let SKUCList(ByVal indexUC As Integer, Value As TC_UnitColumn)
    Set pSKUCList(indexUC) = Value
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Public Sub AddUnitColumn(ByVal tUnitColumn As TC_UnitColumn)

    'increment category total
        'SK
        If tUnitColumn.WIPSectionIndex = 4 Then
            pSKTotal = pSKTotal + 1
        '24K (ignoring Raws section: 2)
        ElseIf tUnitColumn.WIPSectionIndex = 1 Or tUnitColumn.WIPSectionIndex = 3 Then
            pTotal = pTotal + 1
        End If
    
    'redim array
        'SK
        If tUnitColumn.WIPSectionIndex = 4 Then
            ReDim Preserve pSKUCList(1 To pSKTotal) As TC_UnitColumn
        '24K (ignoring Raws section: 2)
        ElseIf tUnitColumn.WIPSectionIndex = 1 Or tUnitColumn.WIPSectionIndex = 3 Then
            ReDim Preserve pUCList(1 To pTotal) As TC_UnitColumn
        End If
    
    'add UC to array
        'SK
        If tUnitColumn.WIPSectionIndex = 4 Then
            Set pSKUCList(pSKTotal) = tUnitColumn
        '24K (ignoring Raws section: 2)
        ElseIf tUnitColumn.WIPSectionIndex = 1 Or tUnitColumn.WIPSectionIndex = 3 Then
            Set pUCList(pTotal) = tUnitColumn
        End If

End Sub
