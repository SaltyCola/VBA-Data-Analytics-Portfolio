VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===================================Double Click SN Finder=================================='

'Double Clicking Cell Links
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    'declare variables
    Dim snStr As String 'Last 5 digits of SN as shown in target cell for searching correct tab
    Dim snColor As Long 'Color of SN cell corresponding to part number
    Dim prtNum As String 'Part number of SN cell corresponding to the current search tab
    Dim c As Range 'range iteration object for searching for correct Serial Number
    Dim searchRow As Double 'search row number for finding correct Serial Number
    Dim tanColor As Long 'color corresponding to part number 5319080
    Dim greyColor As Long 'color corresponding to part number 5319180
    Dim blueColor As Long 'color corresponding to part number 5319280
    Dim purpleColor As Long 'color corresponding to part number 5319380
    Dim orangeColor As Long 'color corresponding to part number 5319480
    
    'initialize color variables
    tanColor = RGB(148, 138, 84)
    greyColor = RGB(128, 128, 128)
    blueColor = RGB(0, 32, 96)
    purpleColor = RGB(112, 48, 160)
    orangeColor = RGB(226, 107, 10)
    
    'Prevent Default Cell Edit Mode in target cells only
    If Not Intersect(Target, Range("B10:R409")) Is Nothing Then
        Cancel = True
                        
        'Check for correct format
        If Not (IsEmpty(Target)) And Not (InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Left(Target.Value, 1)) = 0) And Not (InStr(1, "0123456789", Mid(Target.Value, 2, 1)) = 0) And Not (InStr(1, "0123456789", Mid(Target.Value, 3, 1)) = 0) And Not (InStr(1, "0123456789", Mid(Target.Value, 4, 1)) = 0) And Not (InStr(1, "0123456789", Mid(Target.Value, 5, 1)) = 0) Then
            
            'grab SN
            snStr = Target.Value
            
            'get correct part number based on part color
            snColor = Target.Interior.Color
            If snColor = tanColor Then: prtNum = 5319080
            If snColor = greyColor Then: prtNum = 5319180
            If snColor = blueColor Then: prtNum = 5319280
            If snColor = purpleColor Then: prtNum = 5319380
            If snColor = orangeColor Then: prtNum = 5319480
            
            'turn off screen updating
            Application.ScreenUpdating = False
            
            'activate correct tab
            Worksheets(prtNum).Activate
            
            'search for S/N row in current tab and get searchRow
            For Each c In Worksheets(prtNum).Range("B:B")
                If c.Value = "S/N" Then
                    searchRow = c.Row
                    Exit For
                End If
            Next c
            
            'search through row for Serial Number
            For Each c In Worksheets(prtNum).Range(searchRow & ":" & searchRow)
                If Right(c.Value, 5) = snStr Then
                    Application.Goto Worksheets(prtNum).Cells(c.Row, c.Column), Scroll:=True
                    Exit For
                End If
            Next c
            
            'turn on screen updating
            Application.ScreenUpdating = True
            
        End If

    End If
        
End Sub
