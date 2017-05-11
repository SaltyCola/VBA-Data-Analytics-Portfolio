Attribute VB_Name = "TC_TestSubs"
Public Sub Fix_ChartsWIPSheetReference()

    Dim c As Range
    
    For Each c In ActiveSheet.Range(Cells(170, 3), Cells(198, 3))
        If Left(c.Formula, 6) = "=#REF!" Then
            c.Formula = Replace(c.Formula, "=#REF!", "='NEO 5322121'!")
        End If
    Next c

End Sub

Public Sub Fix_ShippedDates()

    Dim c As Range
    Dim strDate As String
    
    For Each c In ActiveSheet.Range(Cells(7, 3), Cells(7, 436))
        If Not Left(CStr(c.Value), 5) = "Error" And Not IsEmpty(c) Then
            If Len(CStr(c.Value)) > 5 Then
                strDate = Left(CStr(c.Value), (Len(CStr(c.Value)) - 5)) & "/2015"
                c.Value = CDate(strDate)
            End If
        ElseIf Left(CStr(c.Value), 5) = "Error" Then
            c.Value = ""
        End If
    Next c

End Sub

Public Sub test_Variant0Value()

    Dim v As Variant
    
    If v = "" Then: MsgBox v

End Sub

Public Sub test_ForDecrement()

    Dim i As Integer
    
    For i = 1 To 10
        MsgBox i
        i = i - 1
    Next i

End Sub

Public Sub DateComparisonCheck()

    Dim d1 As Date
    Dim d2 As Date
    
    d1 = CDate("7/7/2016")
    d2 = CDate("7/11/2016")
    
    If d1 < d2 Then: MsgBox True

End Sub

Public Sub USEFULL_GrabFormulaOnlyIfFormula()

    Dim c As Range
    
    Set c = Range("C3")
    
    If Left(c.Formula, 1) = "=" Then
        MsgBox c.Formula
    Else
        MsgBox c.Value
    End If

End Sub

Public Sub USEFUL_CheckForComment()

    Dim c As Range
    Dim cmt As Comment
    
    Set c = Range("C3")
    
    If Not c.Comment Is Nothing Then
        cmt = c.Comment
        MsgBox cmt.Text
    Else
        MsgBox "No Comment"
    End If
    
    c.Offset(1, 0).ClearComments
    
    If Not cmt Is Nothing Then
        c.Offset(1, 0).AddComment cmt.Text
    End If
    
End Sub

Public Sub ClassTreeTesting()

    Dim tUC As TC_UnitColumn
    Dim c As Range
    
    Set tUC = New TC_UnitColumn
    Set c = Cells(6, 3)
    
    'assign properties
    tUC.ColumnAddress = c.Address
    tUC.ColumnNumber = c.Column
    tUC.PartNumber = "5322121"
    tUC.TrackingNumber = c.Value
    tUC.TNumAbbr = Right(c.Value, 5)
    tUC.WaterfallIndex = 1
    
    'grab remaining data through methods
    tUC.Headers.GrabData Cells(1, 3), 6
    tUC.GrabOperationsData Cells(7, 3), 42
    tUC.Notes.GrabData Cells(49, 3), 9

    If True Then: 'BreakPoint

End Sub

Public Sub test_CollectionWriting()

    Dim cColl As Collection
    Dim tRng As Range
    
    Set cColl = New Collection
    
    Set tRng = Cells(7, 23)
    
    cColl.Add tRng
    cColl.Add tRng.Offset(23, 56)
    cColl.Add tRng.Offset(87, -3)
    
    Set tRng = cColl(2)
    
    Set tRng = tRng.Offset(-29, -78)
    
    Set cColl(2) = tRng

End Sub

Public Sub CheckForDifferences()

    Dim c As Range
    
    Workbooks("Testing after 1 run").Worksheets("NEO 5322121").Activate
    
    For Each c In Range("A1:VR100")
    
        If Worksheets("after").Cells(c.Row, c.Column).Value <> c.Value Then: MsgBox c.Address
        If Worksheets("after").Cells(c.Row, c.Column).Interior.Color <> c.Interior.Color Then: MsgBox c.Address
        If Not c.Comment Is Nothing Then
            If Not Worksheets("after").Cells(c.Row, c.Column).Comment Is Nothing Then
                If Worksheets("after").Cells(c.Row, c.Column).Comment.Text <> c.Comment.Text Then: MsgBox c.Address
            Else
                MsgBox c.Address
            End If
        End If
        
    Next c

End Sub
