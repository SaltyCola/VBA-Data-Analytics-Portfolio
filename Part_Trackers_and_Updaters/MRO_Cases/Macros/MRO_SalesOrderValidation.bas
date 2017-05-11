Attribute VB_Name = "MRO_SalesOrderValidation"
Sub SalesOrderNumberValidation()

    'declare variables
    Dim salesOrder As String 'variable for holding current sales order number
    Dim sr As Range 'Generic Range Iteration variable for iterating Ship Record tab
    Dim mro As Range 'Generic range iteration variable for iterating MRO tab
    Dim clrOrange As Long 'Color for Sales Order #s not on MRO tab
    Dim clrPurple As Long 'Color for Sales Order #s in MRO tab
    Dim clrGreen As Long 'Color for completed ops
    
    'initialize color variables
    clrOrange = RGB(255, 192, 0)
    clrPurple = RGB(204, 192, 218)
    clrGreen = RGB(146, 208, 80)
    
    'iterate Sales Orders on Ship Record tab
    For Each sr In Worksheets("Ship Record").Range("A:A")
        
        'reached end of list
        If (sr.Row > 2) And (IsEmpty(sr)) Then
            Exit For
        
        'non validated sales order #
        ElseIf (sr.Row > 2) And (sr.Interior.Color = RGB(255, 255, 255)) Then
            
            'grab sales order number
            salesOrder = sr.Value
            
            'iterate sales order row in MRO tab
            For Each mro In Worksheets("MRO").Range("13:13")
                
                'reached end of list before sales order number is found
                If (mro.Column > 2) And (IsEmpty(mro)) Then
                    
                    'make sr row orange
                    Worksheets("Ship Record").Cells(sr.Row, 1).Interior.Color = clrOrange
                    Worksheets("Ship Record").Cells(sr.Row, 2).Interior.Color = clrOrange
                    Worksheets("Ship Record").Cells(sr.Row, 3).Interior.Color = clrOrange
                    
                    'exit sales order search
                    Exit For
                
                'sales order number found
                ElseIf (mro.Column > 2) And (mro.Value = salesOrder) Then
                    
                    'change header
                    Worksheets("MRO").Cells(8, mro.Column).Value = "SHIPPED"
                    
                    'make subheaders purple
                    Worksheets("MRO").Cells(9, mro.Column).Interior.Color = clrPurple
                    Worksheets("MRO").Cells(10, mro.Column).Interior.Color = clrPurple
                    Worksheets("MRO").Cells(11, mro.Column).Interior.Color = clrPurple
                    'Worksheets("MRO").Cells(12, mro.Column).Interior.Color = clrPurple 'SKIPPED BECAUSE COLOR CODED PART NUMBERS
                    Worksheets("MRO").Cells(13, mro.Column).Interior.Color = clrPurple
                    
                    'make ops green
                    Worksheets("MRO").Cells(14, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(15, mro.Column).Interior.Color = clrGreen
                    'Worksheets("MRO").Cells(17, mro.Column).Interior.Color = clrGreen 'SKIPPED BECAUSE HIDDEN
                    Worksheets("MRO").Cells(16, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(18, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(19, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(20, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(21, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(22, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(23, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(24, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(25, mro.Column).Interior.Color = clrGreen
                    Worksheets("MRO").Cells(26, mro.Column).Interior.Color = clrGreen
                    
                    'make sr row purple
                    Worksheets("Ship Record").Cells(sr.Row, 1).Interior.Color = clrPurple
                    Worksheets("Ship Record").Cells(sr.Row, 2).Interior.Color = clrPurple
                    Worksheets("Ship Record").Cells(sr.Row, 3).Interior.Color = clrPurple
                    
                    'exit sales order search
                    Exit For
                    
                End If
                
            Next mro
            
        End If
        
    Next sr

End Sub

