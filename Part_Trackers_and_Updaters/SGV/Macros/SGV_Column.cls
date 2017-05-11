VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SGV_Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================'
'================================== Properties =================================='
'================================================================================'

Private pPrtNum As String 'Contains SN's part number
Private pSerNum As String 'Contains Serial Number Text
Private pQnArray As Variant 'Contains QN cells text in 1st col and colors in 2nd col
Private pOpDateArray As Variant 'Contains Op cells dates in 1st col and colors in 2nd col
Private pNotesArray As Variant 'Contains Notes cells text in 1st col and colors in 2nd col


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Part Number Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get PartNumber() As String
    PartNumber = pPrtNum
End Property
Public Property Let PartNumber(Value As String)
    pPrtNum = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Serial Number Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get SerialNumber() As String
    SerialNumber = pSerNum
End Property
Public Property Let SerialNumber(Value As String)
    pSerNum = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' QNs Array Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get QnArray() As Variant
    QnArray = pQnArray
End Property
Public Property Let QnArray(Value As Variant)
    pQnArray = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Operation Dates Array Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpDateArray() As Variant
    OpDateArray = pOpDateArray
End Property
Public Property Let OpDateArray(Value As Variant)
    pOpDateArray = Value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Notes Array Property
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NotesArray() As Variant
    NotesArray = pNotesArray
End Property
Public Property Let NotesArray(Value As Variant)
    pNotesArray = Value
End Property


'================================================================================'
'==================================== Methods ==================================='
'================================================================================'
