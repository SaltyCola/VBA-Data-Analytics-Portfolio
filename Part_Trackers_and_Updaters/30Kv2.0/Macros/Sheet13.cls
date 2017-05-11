VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===================================Double Click SN Finder=================================='
Option Explicit

'Double Clicking Cell Links
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dim trgtStr As String
    Dim alphnum As Boolean
    Dim snRng As Range
    Dim sn As Range
    Dim snCol As Long
    Dim snRow As Long
    Dim snTxtRaw As String
    Dim snTxt As String
    Dim snTxtPrefix
    Dim snTxtU As String
    Dim snTxtL As String
    
    'Prevent Default Cell Edit Mode
    Cancel = True
    
    'Define snRng
    Set snRng = Worksheets("NEO 5322121").Range("6:6")
    
    'Define trgtStr
    If (Target.Column <= 35) And (Target.Column >= 4) And (Target.Row >= 13) Then
        trgtStr = Target.Value
    End If

    'Call AlphaNumeric Function
    alphnum = AlphaNumeric(trgtStr, Target)
    
    'Cell is AlphaNumeric and therefore Serial Number
    If alphnum Then
        'search sn row in tracker
        For Each sn In snRng
            snTxtRaw = sn.Value
            snRow = sn.Row
            snCol = sn.Column
            'only look at SN's
            If Len(snTxtRaw) >= 10 Then
                snTxt = Right(snTxtRaw, (Len(snTxtRaw) - 5))
                snTxtPrefix = Mid(snTxt, 1, 1)
                snTxtU = UCase(snTxtPrefix) & Mid(snTxt, 2)
                snTxtL = LCase(snTxtPrefix) & Mid(snTxt, 2)
                If (trgtStr = snTxtU) Or (trgtStr = snTxtL) Then
                    Application.Goto Worksheets("NEO 5322121").Cells(snRow, snCol), Scroll:=True
                End If
            End If
        Next sn
    End If
        
End Sub

Function AlphaNumeric(pValue, ByVal Target As Range) As Boolean

    Dim LPos As Double
    Dim LChar As String
    Dim LValid_ValuesA As String
    Dim LValid_ValuesN As String
    
    'start at first char in pValue
    LPos = 1
    
    'Set up alphanumeric values for testing pValue
    LValid_ValuesA = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    LValid_ValuesN = "1234567890"
    
    'Test each char in pValue
    While LPos <= Len(pValue)
        'single character in pValue
        LChar = Mid(pValue, LPos, 1)
        'if char is not alphanumeric, return FALSE
        If (InStr(LValid_ValuesA, LChar) = 0) And (InStr(LValid_ValuesN, LChar) = 0) Then
            AlphaNumeric = False
            Exit Function
        End If
        
        'Increment counter
        LPos = LPos + 1
        
    Wend
    
    'Value is alphanumeric, return TRUE
    AlphaNumeric = True

End Function

