VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGV_EngineSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pSetNum As Integer
Private psp080 As Integer 'column start position
Private psp180 As Integer 'column start position
Private psp280 As Integer 'column start position
Private psp380 As Integer 'column start position
Private psp480 As Integer 'column start position
Private pep080 As Integer 'column end position
Private pep180 As Integer 'column end position
Private pep280 As Integer 'column end position
Private pep380 As Integer 'column end position
Private pep480 As Integer 'column end position
Private parr080 As SGV_EngineSetPNArrays '080 array
Private parr180 As SGV_EngineSetPNArrays '180 array
Private parr280 As SGV_EngineSetPNArrays '280 array
Private parr380 As SGV_EngineSetPNArrays '380 array
Private parr480 As SGV_EngineSetPNArrays '480 array


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Engine Set Number
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SetNum() As Integer
    SetNum = pSetNum
End Property
Public Property Let SetNum(Value As Integer)
    pSetNum = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 080 starting position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get sp080() As Integer
    sp080 = psp080
End Property
Public Property Let sp080(Value As Integer)
    psp080 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 080 ending position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ep080() As Integer
    ep080 = pep080
End Property
Public Property Let ep080(Value As Integer)
    pep080 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 080 Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arr080() As SGV_EngineSetPNArrays
    Set arr080 = parr080
End Property
Public Property Let arr080(Value As SGV_EngineSetPNArrays)
    Set parr080 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 180 starting position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get sp180() As Integer
    sp180 = psp180
End Property
Public Property Let sp180(Value As Integer)
    psp180 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 180 ending position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ep180() As Integer
    ep180 = pep180
End Property
Public Property Let ep180(Value As Integer)
    pep180 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 180 Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arr180() As SGV_EngineSetPNArrays
    Set arr180 = parr180
End Property
Public Property Let arr180(Value As SGV_EngineSetPNArrays)
    Set parr180 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 280 starting position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get sp280() As Integer
    sp280 = psp280
End Property
Public Property Let sp280(Value As Integer)
    psp280 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 280 ending position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ep280() As Integer
    ep280 = pep280
End Property
Public Property Let ep280(Value As Integer)
    pep280 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 280 Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arr280() As SGV_EngineSetPNArrays
    Set arr280 = parr280
End Property
Public Property Let arr280(Value As SGV_EngineSetPNArrays)
    Set parr280 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 380 starting position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get sp380() As Integer
    sp380 = psp380
End Property
Public Property Let sp380(Value As Integer)
    psp380 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 380 ending position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ep380() As Integer
    ep380 = pep380
End Property
Public Property Let ep380(Value As Integer)
    pep380 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 380 Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arr380() As SGV_EngineSetPNArrays
    Set arr380 = parr380
End Property
Public Property Let arr380(Value As SGV_EngineSetPNArrays)
    Set parr380 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 480 starting position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get sp480() As Integer
    sp480 = psp480
End Property
Public Property Let sp480(Value As Integer)
    psp480 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 480 ending position
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ep480() As Integer
    ep480 = pep480
End Property
Public Property Let ep480(Value As Integer)
    pep480 = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' 480 Array
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get arr480() As SGV_EngineSetPNArrays
    Set arr480 = parr480
End Property
Public Property Let arr480(Value As SGV_EngineSetPNArrays)
    Set parr480 = Value
End Property


'==============================================================================='
'=================================== Methods ==================================='
'==============================================================================='


Private Sub Class_Initialize()

    'initialize pn array properties
    Set parr080 = New SGV_EngineSetPNArrays
    Set parr180 = New SGV_EngineSetPNArrays
    Set parr280 = New SGV_EngineSetPNArrays
    Set parr380 = New SGV_EngineSetPNArrays
    Set parr480 = New SGV_EngineSetPNArrays
    
End Sub

Public Sub PrintEngSet()

    Dim c As Range
    Dim printCol As Integer
    Dim finalRow As Integer
    Dim LeadTime As Integer
    
    
    'activate Engine Set Progress Tab
    Worksheets("Engine Set Progress").Activate
    
    
    '080===========================================================================
    printCol = 2 + ((Me.SetNum - 1) * 11)
    finalRow = 0
        'op row 20
        If Me.arr080.Op20 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (20)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op20
                    finalRow = c.Row
                    LeadTime = 0
                    Exit For
                End If
            Next c
        End If
        'op row 21
        If Me.arr080.Op21 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (21)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op21
                    finalRow = c.Row
                    LeadTime = 2
                    Exit For
                End If
            Next c
        End If
        'op row 22
        If Me.arr080.Op22 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (22)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op22
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 23
        If Me.arr080.Op23 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (23)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op23
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 24
        If Me.arr080.Op24 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (24)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op24
                    finalRow = c.Row
                    LeadTime = 13
                    Exit For
                End If
            Next c
        End If
        'op row 25
        If Me.arr080.Op25 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (25)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op25
                    finalRow = c.Row
                    LeadTime = 14
                    Exit For
                End If
            Next c
        End If
        'op row 26
        If Me.arr080.Op26 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (26)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op26
                    finalRow = c.Row
                    LeadTime = 15
                    Exit For
                End If
            Next c
        End If
        'op row 27
        If Me.arr080.Op27 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (27)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op27
                    finalRow = c.Row
                    LeadTime = 17
                    Exit For
                End If
            Next c
        End If
        'op row 28
        If Me.arr080.Op28 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (28)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op28
                    finalRow = c.Row
                    LeadTime = 18
                    Exit For
                End If
            Next c
        End If
        'op row 29
        If Me.arr080.Op29 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (29)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op29
                    finalRow = c.Row
                    LeadTime = 28
                    Exit For
                End If
            Next c
        End If
        'op row 30
        If Me.arr080.Op30 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (30)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op30
                    finalRow = c.Row
                    LeadTime = 29
                    Exit For
                End If
            Next c
        End If
        'op row 31
        If Me.arr080.Op31 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (31)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op31
                    finalRow = c.Row
                    LeadTime = 30
                    Exit For
                End If
            Next c
        End If
        'op row 32
        If Me.arr080.Op32 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (32)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op32
                    finalRow = c.Row
                    LeadTime = 31
                    Exit For
                End If
            Next c
        End If
        'op row 33
        If Me.arr080.Op33 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (33)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op33
                    finalRow = c.Row
                    LeadTime = 32
                    Exit For
                End If
            Next c
        End If
        'op row 34
        If Me.arr080.Op34 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (34)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op34
                    finalRow = c.Row
                    LeadTime = 33
                    Exit For
                End If
            Next c
        End If
        'op row 35
        If Me.arr080.Op35 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr080.OpAbrvGrabber (35)
                    c.Value = Me.arr080.OpAbrv
                    c.Offset(0, 1).Value = Me.arr080.Op35
                    finalRow = c.Row
                    LeadTime = 34
                    Exit For
                End If
            Next c
        End If
        'bold final row
        If finalRow > 0 Then
            Worksheets("Engine Set Progress").Cells(finalRow, printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells(finalRow, (printCol + 1)).Font.Bold = True
            'lead times
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Value = "Lead Time:"
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Value = LeadTime
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Bold = True
        End If
    '080===========================================================================
        
        
    '180===========================================================================
    printCol = printCol + 2
    finalRow = 0
    LeadTime = 0
        'op row 20
        If Me.arr180.Op20 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (20)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op20
                    finalRow = c.Row
                    LeadTime = 0
                    Exit For
                End If
            Next c
        End If
        'op row 21
        If Me.arr180.Op21 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (21)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op21
                    finalRow = c.Row
                    LeadTime = 2
                    Exit For
                End If
            Next c
        End If
        'op row 22
        If Me.arr180.Op22 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (22)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op22
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 23
        If Me.arr180.Op23 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (23)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op23
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 24
        If Me.arr180.Op24 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (24)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op24
                    finalRow = c.Row
                    LeadTime = 13
                    Exit For
                End If
            Next c
        End If
        'op row 25
        If Me.arr180.Op25 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (25)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op25
                    finalRow = c.Row
                    LeadTime = 14
                    Exit For
                End If
            Next c
        End If
        'op row 26
        If Me.arr180.Op26 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (26)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op26
                    finalRow = c.Row
                    LeadTime = 15
                    Exit For
                End If
            Next c
        End If
        'op row 27
        If Me.arr180.Op27 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (27)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op27
                    finalRow = c.Row
                    LeadTime = 17
                    Exit For
                End If
            Next c
        End If
        'op row 28
        If Me.arr180.Op28 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (28)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op28
                    finalRow = c.Row
                    LeadTime = 18
                    Exit For
                End If
            Next c
        End If
        'op row 29
        If Me.arr180.Op29 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (29)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op29
                    finalRow = c.Row
                    LeadTime = 28
                    Exit For
                End If
            Next c
        End If
        'op row 30
        If Me.arr180.Op30 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (30)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op30
                    finalRow = c.Row
                    LeadTime = 29
                    Exit For
                End If
            Next c
        End If
        'op row 31
        If Me.arr180.Op31 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (31)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op31
                    finalRow = c.Row
                    LeadTime = 30
                    Exit For
                End If
            Next c
        End If
        'op row 32
        If Me.arr180.Op32 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (32)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op32
                    finalRow = c.Row
                    LeadTime = 31
                    Exit For
                End If
            Next c
        End If
        'op row 33
        If Me.arr180.Op33 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (33)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op33
                    finalRow = c.Row
                    LeadTime = 32
                    Exit For
                End If
            Next c
        End If
        'op row 34
        If Me.arr180.Op34 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (34)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op34
                    finalRow = c.Row
                    LeadTime = 33
                    Exit For
                End If
            Next c
        End If
        'op row 35
        If Me.arr180.Op35 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr180.OpAbrvGrabber (35)
                    c.Value = Me.arr180.OpAbrv
                    c.Offset(0, 1).Value = Me.arr180.Op35
                    finalRow = c.Row
                    LeadTime = 34
                    Exit For
                End If
            Next c
        End If
        'bold final row
        If finalRow > 0 Then
            Worksheets("Engine Set Progress").Cells(finalRow, printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells(finalRow, (printCol + 1)).Font.Bold = True
            'lead times
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Value = "Lead Time:"
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Value = LeadTime
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Bold = True
        End If
    '180===========================================================================
        
        
    '280===========================================================================
    printCol = printCol + 2
    finalRow = 0
    LeadTime = 0
        'op row 20
        If Me.arr280.Op20 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (20)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op20
                    finalRow = c.Row
                    LeadTime = 0
                    Exit For
                End If
            Next c
        End If
        'op row 21
        If Me.arr280.Op21 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (21)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op21
                    finalRow = c.Row
                    LeadTime = 2
                    Exit For
                End If
            Next c
        End If
        'op row 22
        If Me.arr280.Op22 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (22)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op22
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 23
        If Me.arr280.Op23 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (23)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op23
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 24
        If Me.arr280.Op24 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (24)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op24
                    finalRow = c.Row
                    LeadTime = 13
                    Exit For
                End If
            Next c
        End If
        'op row 25
        If Me.arr280.Op25 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (25)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op25
                    finalRow = c.Row
                    LeadTime = 14
                    Exit For
                End If
            Next c
        End If
        'op row 26
        If Me.arr280.Op26 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (26)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op26
                    finalRow = c.Row
                    LeadTime = 15
                    Exit For
                End If
            Next c
        End If
        'op row 27
        If Me.arr280.Op27 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (27)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op27
                    finalRow = c.Row
                    LeadTime = 17
                    Exit For
                End If
            Next c
        End If
        'op row 28
        If Me.arr280.Op28 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (28)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op28
                    finalRow = c.Row
                    LeadTime = 18
                    Exit For
                End If
            Next c
        End If
        'op row 29
        If Me.arr280.Op29 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (29)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op29
                    finalRow = c.Row
                    LeadTime = 28
                    Exit For
                End If
            Next c
        End If
        'op row 30
        If Me.arr280.Op30 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (30)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op30
                    finalRow = c.Row
                    LeadTime = 29
                    Exit For
                End If
            Next c
        End If
        'op row 31
        If Me.arr280.Op31 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (31)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op31
                    finalRow = c.Row
                    LeadTime = 30
                    Exit For
                End If
            Next c
        End If
        'op row 32
        If Me.arr280.Op32 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (32)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op32
                    finalRow = c.Row
                    LeadTime = 31
                    Exit For
                End If
            Next c
        End If
        'op row 33
        If Me.arr280.Op33 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (33)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op33
                    finalRow = c.Row
                    LeadTime = 32
                    Exit For
                End If
            Next c
        End If
        'op row 34
        If Me.arr280.Op34 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (34)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op34
                    finalRow = c.Row
                    LeadTime = 33
                    Exit For
                End If
            Next c
        End If
        'op row 35
        If Me.arr280.Op35 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr280.OpAbrvGrabber (35)
                    c.Value = Me.arr280.OpAbrv
                    c.Offset(0, 1).Value = Me.arr280.Op35
                    finalRow = c.Row
                    LeadTime = 34
                    Exit For
                End If
            Next c
        End If
        'bold final row
        If finalRow > 0 Then
            Worksheets("Engine Set Progress").Cells(finalRow, printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells(finalRow, (printCol + 1)).Font.Bold = True
            'lead times
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Value = "Lead Time:"
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Value = LeadTime
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Bold = True
        End If
    '280===========================================================================
        
        
    '380===========================================================================
    printCol = printCol + 2
    finalRow = 0
    LeadTime = 0
        'op row 20
        If Me.arr380.Op20 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (20)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op20
                    finalRow = c.Row
                    LeadTime = 0
                    Exit For
                End If
            Next c
        End If
        'op row 21
        If Me.arr380.Op21 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (21)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op21
                    finalRow = c.Row
                    LeadTime = 2
                    Exit For
                End If
            Next c
        End If
        'op row 22
        If Me.arr380.Op22 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (22)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op22
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 23
        If Me.arr380.Op23 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (23)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op23
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 24
        If Me.arr380.Op24 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (24)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op24
                    finalRow = c.Row
                    LeadTime = 13
                    Exit For
                End If
            Next c
        End If
        'op row 25
        If Me.arr380.Op25 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (25)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op25
                    finalRow = c.Row
                    LeadTime = 14
                    Exit For
                End If
            Next c
        End If
        'op row 26
        If Me.arr380.Op26 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (26)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op26
                    finalRow = c.Row
                    LeadTime = 15
                    Exit For
                End If
            Next c
        End If
        'op row 27
        If Me.arr380.Op27 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (27)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op27
                    finalRow = c.Row
                    LeadTime = 17
                    Exit For
                End If
            Next c
        End If
        'op row 28
        If Me.arr380.Op28 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (28)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op28
                    finalRow = c.Row
                    LeadTime = 18
                    Exit For
                End If
            Next c
        End If
        'op row 29
        If Me.arr380.Op29 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (29)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op29
                    finalRow = c.Row
                    LeadTime = 28
                    Exit For
                End If
            Next c
        End If
        'op row 30
        If Me.arr380.Op30 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (30)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op30
                    finalRow = c.Row
                    LeadTime = 29
                    Exit For
                End If
            Next c
        End If
        'op row 31
        If Me.arr380.Op31 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (31)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op31
                    finalRow = c.Row
                    LeadTime = 30
                    Exit For
                End If
            Next c
        End If
        'op row 32
        If Me.arr380.Op32 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (32)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op32
                    finalRow = c.Row
                    LeadTime = 31
                    Exit For
                End If
            Next c
        End If
        'op row 33
        If Me.arr380.Op33 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (33)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op33
                    finalRow = c.Row
                    LeadTime = 32
                    Exit For
                End If
            Next c
        End If
        'op row 34
        If Me.arr380.Op34 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (34)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op34
                    finalRow = c.Row
                    LeadTime = 33
                    Exit For
                End If
            Next c
        End If
        'op row 35
        If Me.arr380.Op35 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr380.OpAbrvGrabber (35)
                    c.Value = Me.arr380.OpAbrv
                    c.Offset(0, 1).Value = Me.arr380.Op35
                    finalRow = c.Row
                    LeadTime = 34
                    Exit For
                End If
            Next c
        End If
        'bold final row
        If finalRow > 0 Then
            Worksheets("Engine Set Progress").Cells(finalRow, printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells(finalRow, (printCol + 1)).Font.Bold = True
            'lead times
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Value = "Lead Time:"
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Value = LeadTime
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Bold = True
        End If
    '380===========================================================================
        
        
    '480===========================================================================
    printCol = printCol + 2
    finalRow = 0
    LeadTime = 0
        'op row 20
        If Me.arr480.Op20 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (20)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op20
                    finalRow = c.Row
                    LeadTime = 0
                    Exit For
                End If
            Next c
        End If
        'op row 21
        If Me.arr480.Op21 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (21)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op21
                    finalRow = c.Row
                    LeadTime = 2
                    Exit For
                End If
            Next c
        End If
        'op row 22
        If Me.arr480.Op22 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (22)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op22
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 23
        If Me.arr480.Op23 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (23)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op23
                    finalRow = c.Row
                    LeadTime = 12
                    Exit For
                End If
            Next c
        End If
        'op row 24
        If Me.arr480.Op24 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (24)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op24
                    finalRow = c.Row
                    LeadTime = 13
                    Exit For
                End If
            Next c
        End If
        'op row 25
        If Me.arr480.Op25 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (25)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op25
                    finalRow = c.Row
                    LeadTime = 14
                    Exit For
                End If
            Next c
        End If
        'op row 26
        If Me.arr480.Op26 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (26)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op26
                    finalRow = c.Row
                    LeadTime = 15
                    Exit For
                End If
            Next c
        End If
        'op row 27
        If Me.arr480.Op27 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (27)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op27
                    finalRow = c.Row
                    LeadTime = 17
                    Exit For
                End If
            Next c
        End If
        'op row 28
        If Me.arr480.Op28 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (28)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op28
                    finalRow = c.Row
                    LeadTime = 18
                    Exit For
                End If
            Next c
        End If
        'op row 29
        If Me.arr480.Op29 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (29)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op29
                    finalRow = c.Row
                    LeadTime = 28
                    Exit For
                End If
            Next c
        End If
        'op row 30
        If Me.arr480.Op30 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (30)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op30
                    finalRow = c.Row
                    LeadTime = 29
                    Exit For
                End If
            Next c
        End If
        'op row 31
        If Me.arr480.Op31 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (31)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op31
                    finalRow = c.Row
                    LeadTime = 30
                    Exit For
                End If
            Next c
        End If
        'op row 32
        If Me.arr480.Op32 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (32)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op32
                    finalRow = c.Row
                    LeadTime = 31
                    Exit For
                End If
            Next c
        End If
        'op row 33
        If Me.arr480.Op33 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (33)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op33
                    finalRow = c.Row
                    LeadTime = 32
                    Exit For
                End If
            Next c
        End If
        'op row 34
        If Me.arr480.Op34 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (34)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op34
                    finalRow = c.Row
                    LeadTime = 33
                    Exit For
                End If
            Next c
        End If
        'op row 35
        If Me.arr480.Op35 <> 0 Then
            For Each c In Worksheets("Engine Set Progress").Range(Cells(3, printCol), Cells(18, printCol))
                If IsEmpty(c) Then
                    Me.arr480.OpAbrvGrabber (35)
                    c.Value = Me.arr480.OpAbrv
                    c.Offset(0, 1).Value = Me.arr480.Op35
                    finalRow = c.Row
                    LeadTime = 34
                    Exit For
                End If
            Next c
        End If
        'bold final row
        If finalRow > 0 Then
            Worksheets("Engine Set Progress").Cells(finalRow, printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells(finalRow, (printCol + 1)).Font.Bold = True
            'lead times
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Value = "Lead Time:"
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Value = LeadTime
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Color = RGB(255, 0, 0)
            Worksheets("Engine Set Progress").Cells((finalRow + 2), printCol).Font.Bold = True
            Worksheets("Engine Set Progress").Cells((finalRow + 2), (printCol + 1)).Font.Bold = True
        End If
    '480===========================================================================

End Sub


