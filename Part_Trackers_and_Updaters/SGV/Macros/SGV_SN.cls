VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGV_SN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pSN As String
Private pOp As String
Private pOpDate As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Serial Number Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SerialNumber() As String
    SerialNumber = pSN
End Property
Public Property Let SerialNumber(Value As String)
    pSN = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Last Operation Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LastOp() As String
    LastOp = pOp
End Property
Public Property Let LastOp(Value As String)
    pOp = Value
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Last Date Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get LastDate() As String
    LastDate = pOpDate
End Property
Public Property Let LastDate(Value As String)
    pOpDate = Value
End Property


'==============================================================================='
'=================================== Methods ==================================='
'==============================================================================='

