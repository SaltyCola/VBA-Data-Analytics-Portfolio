VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_CellColor_30K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pBlank As Long 'READ ONLY
Private pComplete As Long 'READ ONLY
Private pBad As Long 'READ ONLY
Private pRTO As Long 'READ ONLY
Private pOX As Long 'READ ONLY
Private pRCVD As Long 'READ ONLY
Private pInHouse As Long 'READ ONLY
Private pSlowRed As Long 'READ ONLY
Private pNumberOfColors As Integer 'READ ONLY
Private parrColorOrder() As Variant 'READ ONLY array for grabbing color order for waterfalling WIP

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Blank Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Blank() As Long
    Blank = pBlank
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Complete Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Complete() As Long
    Complete = pComplete
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Bad Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Bad() As Long
    Bad = pBad
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' RTO Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get RTO() As Long
    RTO = pRTO
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' OX Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OX() As Long
    OX = pOX
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' RCVD Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get RCVD() As Long
    RCVD = pRCVD
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' InHouse Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get InHouse() As Long
    InHouse = pInHouse
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Slow Red Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SlowRed() As Long
    SlowRed = pSlowRed
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of Colors
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NumberOfColors() As Integer
    NumberOfColors = pNumberOfColors
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Color Order Array
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arrColorOrder() As Variant
    arrColorOrder = parrColorOrder
End Property


'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Private Sub Class_Initialize()

    'blank color (not in array)
    pBlank = RGB(255, 255, 255) 'White
    
    'array colors
    pComplete = RGB(146, 208, 80) 'Green
    pBad = RGB(255, 192, 0) 'Orange
    pRTO = RGB(146, 205, 220) 'Blue
    pOX = RGB(79, 98, 40) 'Dark Green
    pRCVD = RGB(79, 98, 40) 'Dark Green
    pInHouse = RGB(177, 160, 199) 'Purple
    pSlowRed = RGB(255, 0, 0) 'Red
    
    'initialize number of colors
    pNumberOfColors = 7
    
    'initialize array (one less than number of colors because Blank is not included here)
    ReDim parrColorOrder(1 To 2, 1 To 6) As Variant
    parrColorOrder(1, 1) = pComplete
        parrColorOrder(2, 1) = 1
    parrColorOrder(1, 2) = pOX 'no need for pRCVD as they are same color
        parrColorOrder(2, 2) = 2
    parrColorOrder(1, 3) = pInHouse
        parrColorOrder(2, 3) = 3
    parrColorOrder(1, 4) = pRTO
        parrColorOrder(2, 4) = 4
    parrColorOrder(1, 5) = pBad
        parrColorOrder(2, 5) = 5
    parrColorOrder(1, 6) = pSlowRed
        parrColorOrder(2, 6) = 6

End Sub

Public Function Items(ByVal index As Integer)

    Dim rtrnColor As Long 'variable to hold color long to return
    
    If index = 0 Then
        rtrnColor = pBlank
    ElseIf index = 1 Then
        rtrnColor = pComplete
    ElseIf index = 2 Then
        rtrnColor = pOX
    ElseIf index = 3 Then
        rtrnColor = pInHouse
    ElseIf index = 4 Then
        rtrnColor = pRTO
    ElseIf index = 5 Then
        rtrnColor = pBad
    ElseIf index = 6 Then
        rtrnColor = pSlowRed
    End If
    
    'return color value
    Items = rtrnColor

End Function
