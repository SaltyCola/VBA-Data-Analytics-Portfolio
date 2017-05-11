VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_CellColor_24K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pBlank As Long 'READ ONLY
Private pOldComplete As Long 'READ ONLY
Private pNewComplete As Long 'READ ONLY
Private pSKComplete As Long 'READ ONLY
Private pBad As Long 'READ ONLY
Private pRTO As Long 'READ ONLY
Private pOldOX_RCVD As Long 'READ ONLY
Private pNewOX_RCVD As Long 'READ ONLY
Private pSKOX_RCVD As Long 'READ ONLY
Private pInHouse As Long 'READ ONLY
Private pFinalSST As Long 'READ ONLY
Private pTAIRescan As Long 'READ ONLY
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
' Old 24K Complete Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OldComplete() As Long
    OldComplete = pOldComplete
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' New 24K Complete Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NewComplete() As Long
    NewComplete = pNewComplete
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' SK 24K Complete Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SKComplete() As Long
    SKComplete = pSKComplete
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
' Old OX/RCVD Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OldOX_RCVD() As Long
    OldOX_RCVD = pOldOX_RCVD
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' New OX/RCVD Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NewOX_RCVD() As Long
    NewOX_RCVD = pNewOX_RCVD
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' SK OX/RCVD Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SKOX_RCVD() As Long
    SKOX_RCVD = pSKOX_RCVD
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
' Final Ship Subject To ("SST") Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FinalSST() As Long
    FinalSST = pFinalSST
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' TAI Rescan Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get TAIRescan() As Long
    TAIRescan = pTAIRescan
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
    pOldComplete = RGB(146, 208, 80) 'Green
    pNewComplete = RGB(247, 150, 70) 'Orange
    pSKComplete = RGB(218, 150, 148) 'Pink
    pBad = RGB(255, 192, 0) 'Yellow-Orange
    pRTO = RGB(146, 205, 220) 'Blue
    pOldOX_RCVD = RGB(79, 98, 40) 'Dark Green
    pNewOX_RCVD = RGB(226, 107, 10) 'Dark Orange
    pSKOX_RCVD = RGB(150, 54, 52) 'Dark Pink
    pInHouse = RGB(177, 160, 199) 'Purple
    pFinalSST = RGB(112, 48, 160) 'Dark Purple
    pTAIRescan = RGB(0, 176, 240) 'Bright Blue
    
    'initialize number of colors
    pNumberOfColors = 12
    
    'initialize array (one less than number of colors because Blank is not included here)
    ReDim parrColorOrder(1 To 2, 1 To 11) As Variant
    parrColorOrder(1, 1) = pOldComplete
        parrColorOrder(2, 1) = 1
    parrColorOrder(1, 2) = pOldOX_RCVD
        parrColorOrder(2, 2) = 2
    parrColorOrder(1, 3) = pNewComplete
        parrColorOrder(2, 3) = 3
    parrColorOrder(1, 4) = pNewOX_RCVD
        parrColorOrder(2, 4) = 4
    parrColorOrder(1, 5) = pSKComplete
        parrColorOrder(2, 5) = 5
    parrColorOrder(1, 6) = pSKOX_RCVD
        parrColorOrder(2, 6) = 6
    parrColorOrder(1, 7) = pInHouse
        parrColorOrder(2, 7) = 7
    parrColorOrder(1, 8) = pRTO
        parrColorOrder(2, 8) = 8
    parrColorOrder(1, 9) = pBad
        parrColorOrder(2, 9) = 9
    parrColorOrder(1, 10) = pFinalSST
        parrColorOrder(2, 10) = 10
    parrColorOrder(1, 11) = pTAIRescan
        parrColorOrder(2, 11) = 11

End Sub

Public Function Items(ByVal index As Integer)

    Dim rtrnColor As Long 'variable to hold color long to return
    
    If index = 0 Then
        rtrnColor = pBlank
    ElseIf index = 1 Then
        rtrnColor = pOldComplete
    ElseIf index = 2 Then
        rtrnColor = pOldOX_RCVD
    ElseIf index = 3 Then
        rtrnColor = pNewComplete
    ElseIf index = 4 Then
        rtrnColor = pNewOX_RCVD
    ElseIf index = 5 Then
        rtrnColor = pSKComplete
    ElseIf index = 6 Then
        rtrnColor = pSKOX_RCVD
    ElseIf index = 7 Then
        rtrnColor = pInHouse
    ElseIf index = 8 Then
        rtrnColor = pRTO
    ElseIf index = 9 Then
        rtrnColor = pBad
    ElseIf index = 10 Then
        rtrnColor = pFinalSST
    ElseIf index = 11 Then
        rtrnColor = pTAIRescan
    End If
    
    'return color value
    Items = rtrnColor

End Function
