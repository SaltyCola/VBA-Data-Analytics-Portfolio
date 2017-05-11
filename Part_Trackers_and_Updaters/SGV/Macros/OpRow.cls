VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

'Functionality
Private pRowNumber As Integer 'READ ONLY row of OpRow's Title cell
Private pIndex As Integer 'READ ONLY Indexed from final (1) to beginning (inf) ops
Private pEnabled As Boolean 'READ ONLY Disabled implies an unused/hidden row in tracker

'Header Values
Private pTitle As String 'READ ONLY Ops Descriptor Name (located in Tracker File Column "B")
Private pLeadTime As Double 'READ ONLY Ops Lead Time Value (located in Tracker File Column "A")

'Unit Column Individual Values
Private pUCDate As Date 'Unit Col's completion Date for this OpRow
Private pUCColorList As CellColor_SGV 'Unit Col's possible colors for this part number '<==Altered
Private pUCColor As Long 'color chosen from list
Private pUCColorOrderIndex As Integer 'chosen color's order index
Private pUCComment As String 'String value of the comment that may exist on this cell

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Row Number
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get RowNumber() As Integer
    RowNumber = pRowNumber
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operation Index (final to beginning)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get index() As Integer
    index = pIndex
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Visible & Enabled or Hidden & Disabled
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operation Title
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Title() As String
    Title = pTitle
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operation Lead Time
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LeadTime() As Double
    LeadTime = pLeadTime
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column DateCell Date
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCDate() As Date
    UCDate = pUCDate
End Property
Public Property Let UCDate(Value As Date)
    pUCDate = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column DateCell Possible Colors List
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCColorList() As CellColor_SGV '<==Altered
    Set UCColorList = pUCColorList
End Property
Public Property Let UCColorList(Value As CellColor_SGV) '<==Altered
    Set pUCColorList = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column DateCell Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCColor() As Long
    UCColor = pUCColor
End Property
Public Property Let UCColor(Value As Long)
    pUCColor = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column DateCell Color Order Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCColorOrderIndex() As Integer
    UCColorOrderIndex = pUCColorOrderIndex
End Property
Public Property Let UCColorOrderIndex(Value As Integer)
    pUCColorOrderIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column DateCell Date
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCComment() As String
    UCComment = pUCComment
End Property
Public Property Let UCComment(Value As String)
    pUCComment = Value
End Property


'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Private Sub Class_Initialize()
    
    'initialize child classes
    Set pUCColorList = New CellColor_SGV '<==Altered

End Sub

Public Sub GrabRowData(ByVal OpTitleRow As Integer, ByVal OpTitleColumn As Integer, ByVal OpRowIndex As Integer, ByVal OpRowVisible As Boolean, ByVal UCColumnNumber As Double)

    Dim i As Integer 'iteration integer for grabbing Unit Column's OpRow Color
    
    pRowNumber = OpTitleRow
    pIndex = OpRowIndex
    pEnabled = OpRowVisible
    
    pTitle = Cells(OpTitleRow, OpTitleColumn).Value
    pLeadTime = Cells(OpTitleRow, (OpTitleColumn - 1)).Value '<==Altered
    
    'Unit Col's OpRow Comment
    If Not Cells(OpTitleRow, UCColumnNumber).Comment Is Nothing Then
        pUCComment = Cells(OpTitleRow, UCColumnNumber).Comment.Text
    Else
        pUCComment = ""
    End If
    
    'Unit Col's OpRow Date
        pUCDate = Cells(OpTitleRow, UCColumnNumber).Value

    'Unit Col's OpRow Color
        'initialize Color to white
        pUCColor = pUCColorList.Items(0)
        'grab color from cell to color list comparison
        For i = 0 To pUCColorList.NumberOfColors
            If pUCColorList.Items(i) = Cells(OpTitleRow, UCColumnNumber).Interior.Color Then
                pUCColor = pUCColorList.Items(i)
                Exit For
            End If
        Next i
        'grab color order index
        For i = 1 To UBound(pUCColorList.arrColorOrder, 2)
            If pUCColor = pUCColorList.arrColorOrder(1, i) Then
                pUCColorOrderIndex = pUCColorList.arrColorOrder(2, i)
                Exit For
            End If
        Next i

End Sub
