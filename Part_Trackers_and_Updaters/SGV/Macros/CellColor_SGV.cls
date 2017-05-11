VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellColor_SGV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pBlank As Long 'READ ONLY
Private pComplete As Long 'READ ONLY
Private pRepairCell As Long 'READ ONLY
Private pBackFromOutsource As Long 'READ ONLY
Private pIHCOutsource As Long 'READ ONLY
Private pFountainOutsource As Long 'READ ONLY
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
' Repair Cell Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get RepairCell() As Long
    RepairCell = pRepairCell
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Back From Outsource Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get BackFromOutsource() As Long
    BackFromOutsource = pBackFromOutsource
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' IHC Outsource Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get IHCOutsource() As Long
    IHCOutsource = pIHCOutsource
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Fountain Plating Outsource Color
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get FountainOutsource() As Long
    FountainOutsource = pFountainOutsource
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
    pRepairCell = RGB(128, 100, 162) 'Purple
    pBackFromOutsource = RGB(250, 191, 143) 'Salmon
    pIHCOutsource = RGB(247, 150, 70) 'Orange
    pFountainOutsource = RGB(255, 255, 0) 'Yellow
    
    'initialize number of colors
    pNumberOfColors = 5
    
    'initialize array
    ReDim parrColorOrder(1 To 2, 1 To 5) As Variant
    parrColorOrder(1, 1) = pComplete
        parrColorOrder(2, 1) = 1
    parrColorOrder(1, 2) = pRepairCell
        parrColorOrder(2, 2) = 2
    parrColorOrder(1, 3) = pBackFromOutsource
        parrColorOrder(2, 3) = 3
    parrColorOrder(1, 4) = pIHCOutsource
        parrColorOrder(2, 4) = 4
    parrColorOrder(1, 5) = pFountainOutsource
        parrColorOrder(2, 5) = 5

End Sub

Public Function Items(ByVal index As Integer)

    Dim rtrnColor As Long 'variable to hold color long to return
    
    If index = 0 Then
        rtrnColor = pBlank
    ElseIf index = 1 Then
        rtrnColor = pComplete
    ElseIf index = 2 Then
        rtrnColor = pRepairCell
    ElseIf index = 3 Then
        rtrnColor = pBackFromOutsource
    ElseIf index = 4 Then
        rtrnColor = pIHCOutsource
    ElseIf index = 5 Then
        rtrnColor = pFountainOutsource
    End If
    
    'return color value
    Items = rtrnColor

End Function
