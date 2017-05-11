VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryColumn_SGV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pSGVPartNumIndex As Integer 'holds the part number row index of the SGV (1 to 5)
Private pOpIndex As Integer 'holds the operation index of the WIPs Operation Row (1 to 16)
Private pOpColumnIndex As Integer 'holds the operation index of the Summary Column (1 to 17)
Private pOpColorOrderIndex As Integer
Private pSGVColor As Long 'holds the color of all summary entries that belong to this object SGV
Private pTotal As Double 'holds the total number of entries in the list
Private pUCList() As Variant 'holds the list of UC's and the corresponding date of their last completed of

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Part Number
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SGVPartNumIndex() As Integer
    SGVPartNumIndex = pSGVPartNumIndex
End Property
Public Property Let SGVPartNumIndex(Value As Integer)
    pSGVPartNumIndex = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Current Op Index
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpIndex() As Integer
    OpIndex = pOpIndex
End Property
Public Property Let OpIndex(Value As Integer)
    pOpIndex = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Column Index
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpColumnIndex() As Integer
    OpColumnIndex = pOpColumnIndex
End Property
Public Property Let OpColumnIndex(Value As Integer)
    pOpColumnIndex = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Color Order Index
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpColorOrderIndex() As Integer
    OpColorOrderIndex = pOpColorOrderIndex
End Property
Public Property Let OpColorOrderIndex(Value As Integer)
    pOpColorOrderIndex = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' SGV Color
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SGVColor() As Long
    SGVColor = pSGVColor
End Property
Public Property Let SGVColor(Value As Long)
    pSGVColor = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Total Number of Unit Columns
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Total() As Double
    Total = pTotal
End Property
Public Property Let Total(Value As Double)
    pTotal = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Unit Column List Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get UCList(ByVal iSNorDate As Integer, ByVal iUCNumber As Integer) As Variant
    UCList = pUCList(iSNorDate, iUCNumber)
End Property
Public Property Let UCList(ByVal iSNorDate As Integer, ByVal iUCNumber As Integer, Value As Variant)
    pUCList(iSNorDate, iUCNumber) = Value
End Property


'=================================================================================='
'==================================== Methods ====================================='
'=================================================================================='

Public Sub InitializeSC()
'must have SGV part number index and op column index set before running this method

    'assign OpIndex
    If pOpColumnIndex = 1 Then
        pOpIndex = 0
    ElseIf pOpColumnIndex <= 9 Then
        pOpIndex = pOpColumnIndex
    ElseIf pOpColumnIndex > 9 Then
        pOpIndex = pOpColumnIndex - 1
    End If
    
    'grab SGV Color
    If pSGVPartNumIndex = 1 Then
        pSGVColor = RGB(148, 138, 84) 'baige
    ElseIf pSGVPartNumIndex = 2 Then
        pSGVColor = RGB(128, 128, 128) 'grey
    ElseIf pSGVPartNumIndex = 3 Then
        pSGVColor = RGB(0, 32, 96) 'dark blue
    ElseIf pSGVPartNumIndex = 4 Then
        pSGVColor = RGB(112, 48, 160) 'purple
    ElseIf pSGVPartNumIndex = 5 Then
        pSGVColor = RGB(226, 107, 10) 'orange
    End If
    
    'initialize color order index
    pOpColorOrderIndex = 1
    
    'repair cell COI
    If pOpColumnIndex = 1 Then: pOpColorOrderIndex = 2
    
    'Fountain COI
    If pOpColumnIndex = 9 Then: pOpColorOrderIndex = 3
    
    'IHC COI
    If pOpColumnIndex = 10 Then: pOpColorOrderIndex = 4

End Sub

Public Sub AddListEntry(ByVal strSerialNum As String, ByVal dateLastOp As Date)

    'increment total
    pTotal = pTotal + 1
    
    'resize array
    ReDim Preserve pUCList(1 To 2, 1 To pTotal) As Variant
    
    'add values to array
    pUCList(1, pTotal) = strSerialNum
    pUCList(2, pTotal) = dateLastOp

End Sub
