Attribute VB_Name = "Mod_30KCheck"
Public Sub QCLiveTrackerCheck()

    Dim c1 As Range
    Dim c2 As Range
    Dim c3 As Range
    Dim c4 As Range
    Dim i As Integer
    Dim arrInfo(1 To 25) As String
    Dim LoadingBar As MsgWaterfall
    Dim SrcFileError As MsgSrcError
    Dim cntWIP As Double
    Dim j As Double
    Dim wrkbWIP As Workbook
    Dim wrkbQC As Workbook
    Dim srcWIP As Boolean
    Dim srcQC As Boolean
    
    
    'initialize booleans
    srcWIP = False
    srcQC = False
    
    
    'check for necessary files openned
    For Each wrkb In Application.Workbooks
        If wrkb.Name = "30K Quality Clinic Live Tracker.xlsm" Then
            Set wrkbQC = wrkb
            srcQC = True
            Exit For
        End If
    Next wrkb
    For Each wrkb In Application.Workbooks
        If Right(wrkb.Name, 29) = "WORKING NEO WIP tracking.xlsm" Then
            Set wrkbWIP = wrkb
            srcWIP = True
            Exit For
        End If
    Next wrkb
    
    
    'source file error messages
    If (Not srcWIP) Then
        Set SrcFileError = New MsgSrcError
        SrcFileError.Show
    End If
    If (Not srcQC) Then
        Set SrcFileError = New MsgSrcError
        SrcFileError.Label1.Caption = "The QC source file is not open. Please open before running the Serial Number Check."
        SrcFileError.Show
    End If
    
    
    'exit sub if source files not open
    If (Not srcWIP) Or (Not srcQC) Then: Exit Sub
    
    
    'clear contents and turn of screen updating
    Range("B2:Z1000").ClearContents
    Application.ScreenUpdating = False
    
    
    'initialize loading bar
    Set LoadingBar = New MsgWaterfall
    LoadingBar.Label1.Caption = "Comparing Files..."
    LoadingBar.Image0.Visible = True
    LoadingBar.Show vbModeless
    DoEvents
    
    
    'count WIP
    cntWIP = 0
    j = 0
    For Each c1 In wrkbWIP.Worksheets("NEO 5322121").Range("6:6")
        'exit for
        If c1.Interior.Color = RGB(255, 0, 0) Then
            Exit For
        ElseIf c1.Column > 2 Then
            cntWIP = cntWIP + 1
        End If
    Next c1
    
    
    'iterate WIP
    For Each c1 In wrkbWIP.Worksheets("NEO 5322121").Range("6:6")
    
        'exit for
        If c1.Interior.Color = RGB(255, 0, 0) Then
            Exit For
        
        'SN cell
        ElseIf c1.Column > 2 Then
        
            'update loading bar
            j = j + 1
            For i = 0 To 166
                LoadingBar.Controls("Image" & i).Visible = False
            Next i
            i = (j / cntWIP) * 166
            LoadingBar.Controls("Image" & i).Visible = True
            LoadingBar.Show vbModeless
            DoEvents
            
            'iterate QC Live Tracker
            For Each c2 In wrkbQC.Worksheets("Quest Tracker").Range("A1:A99999")
            
                'exit for
                If IsEmpty(c2) Then
                    Exit For
                    
                'SN exists in QC Live Tracker
                ElseIf c1.Value = c2.Value Then
                    
                    'find pasting location
                    For Each c3 In ActiveSheet.Range("B2:B999")
                        
                        'paste info
                        If c3.Value = "" Then
                            Range(c2, c2.Offset(0, 24)).Copy Range(c3, c3.Offset(0, 24))
                            'exit for
                            Exit For
                        End If
                        
                    Next c3
                    'exit for
                    Exit For
                    
                End If
            
            Next c2
        
        End If
    
    Next c1
    
    
    'turn on screen updating
    Application.ScreenUpdating = False
    LoadingBar.Hide
    Cells.FormatConditions.Delete
    
    
    'add duplicate values format condition
    Columns("B:B").Select
    Range("B406").Activate
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    

End Sub
