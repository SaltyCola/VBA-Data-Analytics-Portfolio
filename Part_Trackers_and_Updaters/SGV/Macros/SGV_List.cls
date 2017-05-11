VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGV_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pPartNum As String
Private pOp As String
Private pColor As Long
Private pTotal As Double
Private pSGVArray As Variant

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Part Number Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get PartNum() As String
    PartNum = pPartNum
End Property
Public Property Let PartNum(Value As String)
    pPartNum = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Current Op Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Op() As String
    Op = pOp
End Property
Public Property Let Op(Value As String)
    pOp = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Color Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Color() As Long
    Color = pColor
End Property
Public Property Let Color(Value As Long)
    pColor = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Total Property
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
' SGV Array Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SGVArray() As Variant
    SGVArray = pSGVArray
End Property
Public Property Let SGVArray(Value As Variant)
    pSGVArray = Value
End Property


'==============================================================================='
'=================================== Methods ==================================='
'==============================================================================='

Public Sub PrintTotal()

    Dim c As Range
    Dim rw As Double
    Dim col As Double
    
    'activate WIP Summary Tab
    Worksheets("WIP Summary").Activate
    Application.ScreenUpdating = True
    
    'grab Row
    For Each c In Worksheets("WIP Summary").Range("A:A")
        If c.Value = pPartNum Then
            rw = c.Row
            Exit For
        End If
    Next c
    
    'grab col
    For Each c In Worksheets("WIP Summary").Range("4:4")
        If c.Value = pOp Then
            col = c.Column
            Exit For
        End If
    Next c

    'print Total to cell
    If pTotal > 0 Then
        Worksheets("WIP Summary").Cells(rw, col).Value = pTotal
        Worksheets("WIP Summary").Cells(rw, col).Interior.Color = pColor
    'blank out cell if total = 0
    Else
        Worksheets("WIP Summary").Cells(rw, col).Value = ""
        Worksheets("WIP Summary").Cells(rw, col).Interior.Color = RGB(255, 255, 255)
    End If

End Sub

Public Sub PrintList()
    
    Dim c As Range
    Dim rw As Double
    Dim col As Double
    
    'activate WIP Summary Tab
    Worksheets("WIP Summary").Activate
    Application.ScreenUpdating = True
    
    'grab col
    For Each c In Worksheets("WIP Summary").Range("4:4")
        If c.Value = pOp Then
            col = c.Column
            Exit For
        End If
    Next c
    
    'find Row
    For Each c In Worksheets("WIP Summary").Range(Cells(10, col), Cells(310, col))
        If IsEmpty(c) Then
            rw = c.Row
            Exit For
        End If
    Next c
    
    'iterate through array
    For i = 1 To UBound(pSGVArray)
        Worksheets("WIP Summary").Cells((rw + (i - 1)), col).Value = pSGVArray(i)
        Worksheets("WIP Summary").Cells((rw + (i - 1)), col).Interior.Color = pColor
    Next i

End Sub
