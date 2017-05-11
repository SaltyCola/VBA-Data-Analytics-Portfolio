VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGV_EngineSetPNArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pOpAbrv As String
Private pOp20 As Integer
Private pOp21 As Integer
Private pOp22 As Integer
Private pOp23 As Integer
Private pOp24 As Integer
Private pOp25 As Integer
Private pOp26 As Integer
Private pOp27 As Integer
Private pOp28 As Integer
Private pOp29 As Integer
Private pOp30 As Integer
Private pOp31 As Integer
Private pOp32 As Integer
Private pOp33 As Integer
Private pOp34 As Integer
Private pOp35 As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Abbreviation (read only)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpAbrv() As String
    OpAbrv = pOpAbrv
End Property
Public Property Let OpAbrv(Value As String)
    pOpAbrv = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 20
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op20() As Integer
    Op20 = pOp20
End Property
Public Property Let Op20(Value As Integer)
    pOp20 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 21
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op21() As Integer
    Op21 = pOp21
End Property
Public Property Let Op21(Value As Integer)
    pOp21 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 22
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op22() As Integer
    Op22 = pOp22
End Property
Public Property Let Op22(Value As Integer)
    pOp22 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 23
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op23() As Integer
    Op23 = pOp23
End Property
Public Property Let Op23(Value As Integer)
    pOp23 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 24
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op24() As Integer
    Op24 = pOp24
End Property
Public Property Let Op24(Value As Integer)
    pOp24 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 25
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op25() As Integer
    Op25 = pOp25
End Property
Public Property Let Op25(Value As Integer)
    pOp25 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 26
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op26() As Integer
    Op26 = pOp26
End Property
Public Property Let Op26(Value As Integer)
    pOp26 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 27
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op27() As Integer
    Op27 = pOp27
End Property
Public Property Let Op27(Value As Integer)
    pOp27 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 28
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op28() As Integer
    Op28 = pOp28
End Property
Public Property Let Op28(Value As Integer)
    pOp28 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 29
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op29() As Integer
    Op29 = pOp29
End Property
Public Property Let Op29(Value As Integer)
    pOp29 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 30
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op30() As Integer
    Op30 = pOp30
End Property
Public Property Let Op30(Value As Integer)
    pOp30 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 31
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op31() As Integer
    Op31 = pOp31
End Property
Public Property Let Op31(Value As Integer)
    pOp31 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 32
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op32() As Integer
    Op32 = pOp32
End Property
Public Property Let Op32(Value As Integer)
    pOp32 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 33
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op33() As Integer
    Op33 = pOp33
End Property
Public Property Let Op33(Value As Integer)
    pOp33 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 34
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op34() As Integer
    Op34 = pOp34
End Property
Public Property Let Op34(Value As Integer)
    pOp34 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Op Row 35
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op35() As Integer
    Op35 = pOp35
End Property
Public Property Let Op35(Value As Integer)
    pOp35 = Value
End Property


'==============================================================================='
'=================================== Methods ==================================='
'==============================================================================='

Private Sub Class_Initialize()

    'initialize op row counts to zero
    pOp20 = 0
    pOp21 = 0
    pOp22 = 0
    pOp23 = 0
    pOp24 = 0
    pOp25 = 0
    pOp26 = 0
    pOp27 = 0
    pOp28 = 0
    pOp29 = 0
    pOp30 = 0
    pOp31 = 0
    pOp32 = 0
    pOp33 = 0
    pOp34 = 0
    pOp35 = 0

End Sub

Public Sub OpAbrvGrabber(ByVal j As Integer)

    If j = 20 Then
        pOpAbrv = "Shipped"
    ElseIf j = 21 Then
        pOpAbrv = "FX"
    ElseIf j = 22 Then
        pOpAbrv = "QN Submit"
    ElseIf j = 23 Then
        pOpAbrv = "Finish Ops"
    ElseIf j = 24 Then
        pOpAbrv = "Seals"
    ElseIf j = 25 Then
        pOpAbrv = "BL"
    ElseIf j = 26 Then
        pOpAbrv = "Prime"
    ElseIf j = 27 Then
        pOpAbrv = "RX"
    ElseIf j = 28 Then
        pOpAbrv = "Outsource"
    ElseIf j = 29 Then
        pOpAbrv = "OX"
    ElseIf j = 30 Then
        pOpAbrv = "Laser"
    ElseIf j = 31 Then
        pOpAbrv = "Gap Fill"
    ElseIf j = 32 Then
        pOpAbrv = "Cover Bond"
    ElseIf j = 33 Then
        pOpAbrv = "PAA"
    ElseIf j = 34 Then
        pOpAbrv = "Part Mark"
    ElseIf j = 35 Then
        pOpAbrv = "Launch"
    
    ElseIf (i < 20) Or (i > 35) Then
        pOpAbrv = "OP# ERROR"
    
    End If

End Sub

Public Sub OpRowIncrement(ByVal j As Integer)

    Dim tempCount As Integer
    
    If j = 20 Then
        tempCount = pOp20
        tempCount = tempCount + 1
        pOp20 = tempCount
    ElseIf j = 21 Then
        tempCount = pOp21
        tempCount = tempCount + 1
        pOp21 = tempCount
    ElseIf j = 22 Then
        tempCount = pOp22
        tempCount = tempCount + 1
        pOp22 = tempCount
    ElseIf j = 23 Then
        tempCount = pOp23
        tempCount = tempCount + 1
        pOp23 = tempCount
    ElseIf j = 24 Then
        tempCount = pOp24
        tempCount = tempCount + 1
        pOp24 = tempCount
    ElseIf j = 25 Then
        tempCount = pOp25
        tempCount = tempCount + 1
        pOp25 = tempCount
    ElseIf j = 26 Then
        tempCount = pOp26
        tempCount = tempCount + 1
        pOp26 = tempCount
    ElseIf j = 27 Then
        tempCount = pOp27
        tempCount = tempCount + 1
        pOp27 = tempCount
    ElseIf j = 28 Then
        tempCount = pOp28
        tempCount = tempCount + 1
        pOp28 = tempCount
    ElseIf j = 29 Then
        tempCount = pOp29
        tempCount = tempCount + 1
        pOp29 = tempCount
    ElseIf j = 30 Then
        tempCount = pOp30
        tempCount = tempCount + 1
        pOp30 = tempCount
    ElseIf j = 31 Then
        tempCount = pOp31
        tempCount = tempCount + 1
        pOp31 = tempCount
    ElseIf j = 32 Then
        tempCount = pOp32
        tempCount = tempCount + 1
        pOp32 = tempCount
    ElseIf j = 33 Then
        tempCount = pOp33
        tempCount = tempCount + 1
        pOp33 = tempCount
    ElseIf j = 34 Then
        tempCount = pOp34
        tempCount = tempCount + 1
        pOp34 = tempCount
    ElseIf j = 35 Then
        tempCount = pOp35
        tempCount = tempCount + 1
        pOp35 = tempCount
    End If

End Sub


