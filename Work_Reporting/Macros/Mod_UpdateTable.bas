Attribute VB_Name = "Mod_UpdateTable"
Option Explicit

Public Sub UpdateTable()

    Dim c1 As Range 'generic range iteration object
    Dim c2 As Range 'generic range iteration object
    Dim i As Integer 'counter
    Dim j As Integer 'counter
    Dim boolNonMember As Boolean 'boolean for adding new name entries to collection
    Dim arTD(1 To 2, 1 To 28) As Variant 'array of training data by name
    Dim cTD As Collection 'collection of arTD's
    
    'initialize variables
    i = 1
    Set cTD = Nothing
    Set cTD = New Collection
    
    'activate data table column
    Worksheets("Training Matrix").Activate
    
    'sort names in data destination tab
    Call SortNamesDataDestination
    
    'iterate names column
    For Each c1 In Range("A:A")
    
        'initialize arTD
        For j = 1 To 28
            If j < 5 Then
                arTD(1, j) = "N/A"
                arTD(2, j) = RGB(255, 255, 255)
            Else
                arTD(1, j) = 0
                arTD(2, j) = RGB(255, 0, 0)
            End If
        Next j
    
        'end iteration
        If (c1.Row > 7) And (c1.Interior.Color = RGB(0, 0, 0)) Then
            Exit For
        
        ElseIf c1.Row > 7 Then
            
            'iterate through row
            For Each c2 In Range(Cells(c1.Row, 1), Cells(c1.Row, 28))
                arTD(1, c2.Column) = c2.Value
                arTD(2, c2.Column) = c2.Interior.Color
                'update matrix colors
                If c2.Column > 4 Then
                    If c2.Value = 0 Then: arTD(2, c2.Column) = RGB(255, 0, 0)
                    If c2.Value = 1 Then: arTD(2, c2.Column) = RGB(255, 192, 0)
                    If c2.Value = 2 Then: arTD(2, c2.Column) = RGB(255, 255, 0)
                    If c2.Value = 3 Then: arTD(2, c2.Column) = RGB(146, 208, 80)
                    If c2.Value = 4 Then: arTD(2, c2.Column) = RGB(0, 176, 80)
                End If
            Next c2
            
            'add arTD to cTD
            cTD.Add arTD, Str(i)
            i = i + 1
        
        End If
    
    Next c1
    
    'activate data source tab
    Worksheets("Direct Reports").Activate
    
    'Sort Names in data source tab
    Call SortNamesDataSource
    
    'reset counter
    i = 1
    
    'iterate through Names Column
    For Each c1 In Range("B:B")
    
        'end iteration
        If c1.Interior.Color = RGB(0, 0, 0) Then
            Exit For
        
        'skip row 1
        ElseIf c1.Row > 1 Then
        
            'initialize boolean
            boolNonMember = True
            
            'iterate cTD for name membership
            For i = 1 To cTD.Count
            
                'test for membership
                If cTD(i)(1, 1) = c1.Value Then
                
                    'update values
                    cTD(i)(1, 2) = Cells(c1.Row, 9).Value
                    cTD(i)(1, 3) = Cells(c1.Row, 5).Value
                    cTD(i)(1, 4) = Cells(c1.Row, 10).Value
                    
                    'set boolean
                    boolNonMember = False
                
                End If
            
            Next i
            
            'name is not a member of collection
            If boolNonMember Then
            
                'initialize arTD
                For i = 1 To 28
                    If i < 5 Then
                        arTD(1, i) = "N/A"
                        arTD(2, i) = RGB(255, 255, 255)
                    Else
                        arTD(1, i) = 0
                        arTD(2, i) = RGB(255, 0, 0)
                    End If
                Next i
                
                'grab data
                arTD(1, 1) = c1.Value
                arTD(1, 2) = Cells(c1.Row, 9).Value
                arTD(1, 3) = Cells(c1.Row, 5).Value
                arTD(1, 4) = Cells(c1.Row, 10).Value
                
                'add to end of collection
                i = cTD.Count + 1
                cTD.Add arTD, Str(i)
            
            End If
        
        End If
    
    Next c1
    
    'activate data destination tab
    Worksheets("Training Matrix").Activate
    
    'find final (black) row
    For Each c1 In Range("A:A")
        If (c1.Row > 7) And (c1.Interior.Color = RGB(0, 0, 0)) Then
            i = c1.Row
            Exit For
        End If
    Next c1
    
    'delete table except for black bar
    Range(Cells(8, 1), Cells((i - 1), 28)).Delete
    
    'disable screen updates
    'Application.ScreenUpdating = False
    
    'insert lines for printing
    Rows(8 & ":" & (8 + (cTD.Count - 1))).Insert
    Rows(8 & ":" & (8 + (cTD.Count - 1))).Interior.Color = RGB(255, 255, 255)
    Rows(8 & ":" & (8 + (cTD.Count - 1))).RowHeight = 15
    Rows(8 & ":" & (8 + (cTD.Count - 1))).Font.Bold = False
    Rows(8 & ":" & (8 + (cTD.Count - 1))).Font.Italic = False
    
    'print collection to table
    For Each c1 In Range("A:A")
    
        'end iteration
        If (c1.Row > 7) And (c1.Interior.Color = RGB(0, 0, 0)) Then
            Exit For
        
        'add info
        ElseIf (c1.Row > 7) And Not (c1.Interior.Color = RGB(0, 0, 0)) Then
            For Each c2 In Range(Cells(c1.Row, 1), Cells(c1.Row, 28))
                c2.Value = cTD(c1.Row - 7)(1, c2.Column)
                c2.Interior.Color = cTD(c1.Row - 7)(2, c2.Column)
            Next c2
        
        End If
    
    Next c1
    
    'resort data table
    Call SortNamesDataDestination
    
    'empty cTD
    Set cTD = Nothing
    
    'hyperlinks
    Call AddHyperlinks
    
    'enable screen updates
    Application.ScreenUpdating = True

End Sub

Public Sub SortNamesDataSource()

    ActiveWorkbook.Worksheets("Direct Reports").ListObjects("Table1").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Direct Reports").ListObjects("Table1").Sort. _
        SortFields.Add Key:=Range("Table1[[#All],[Name]]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Direct Reports").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Public Sub SortNamesDataDestination()

    ActiveWorkbook.Worksheets("Training Matrix").ListObjects("Table256").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Training Matrix").ListObjects("Table256").Sort. _
        SortFields.Add Key:=Range("Table256[[#All],[Name]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Training Matrix").ListObjects("Table256"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Public Sub AddHyperlinks()

    Dim c As Range
    
    'activate data destination
    Worksheets("Training Matrix").Activate
    
    'iterate Names
    For Each c In Range("A:A")
    
        If (c.Row > 7) And (c.Interior.Color = RGB(0, 0, 0)) Then
            Exit For
        
        ElseIf (c.Row > 7) And Not (c.Interior.Color = RGB(0, 0, 0)) Then
            'check for folder existance
            If Len(Dir(ActiveWorkbook.Path & "\Training\Training Records\" & c.Value, vbDirectory)) = 0 Then
                'create directory
                'MkDir ActiveWorkbook.Path & "\Training\Training Records\" & c.Value
            Else
                'place hyperlink
                c.Hyperlinks.Add Anchor:=c, Address:="Training\Training Records\" & c.Value, TextToDisplay:=c.Value
            End If
        
        End If
    
    Next c

End Sub



