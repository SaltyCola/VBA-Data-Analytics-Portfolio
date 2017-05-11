Attribute VB_Name = "Mod_CondColorFormats"
Option Explicit

Public Sub WhiteCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("A5").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub GreenCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("A3").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub LightGreenCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("A4").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub LightYellowCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("B4").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub YellowCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("B5").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub DarkYellowCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("B6").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub RedCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("A8").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub OrangeCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("A6").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub BlueCell()

    ActiveCell.Interior.Color = Worksheets("List Source").Range("B7").Interior.Color
    Call DataBars
    Call SNCheck

End Sub

Public Sub SNCheck()

    Dim blckcll As Range
    Dim FinalRow As Double
    Dim del As Range
    Dim SNDuplicate As FormatCondition
    
    'initialize variables
    FinalRow = 0
    
    Application.ScreenUpdating = False

    'find final row
    For Each blckcll In ActiveSheet.Range("A:A")
        If blckcll.Interior.Color = RGB(9, 10, 11) Then
            FinalRow = blckcll.Row - 1
            Exit For
        End If
    Next blckcll
    
    'delete all conditional formatting
    For Each del In ActiveSheet.Range(Cells(1, 3), Cells(FinalRow, 3))
        del.FormatConditions.Delete
    Next del
    
    Application.ScreenUpdating = True
    
    'create conditional formatting
    ActiveSheet.Range(Cells(6, 3), Cells(FinalRow, 3)).FormatConditions.AddUniqueValues
    ActiveSheet.Range(Cells(6, 3), Cells(FinalRow, 3)).FormatConditions(1).DupeUnique = xlDuplicate
    ActiveSheet.Range(Cells(6, 3), Cells(FinalRow, 3)).FormatConditions(1).Interior.Color = RGB(218, 150, 148)

End Sub

Public Sub DataBars()

    Dim blckcll As Range
    Dim FinalRow As Double
    Dim del As Range
    Dim CellDataBar As Databar
    Dim CellDataBarColor As FormatColor
    
    'initialize variables
    FinalRow = 0
    
    Application.ScreenUpdating = False
    
    'find final row
    For Each blckcll In ActiveSheet.Range("A:A")
        If blckcll.Interior.Color = RGB(9, 10, 11) Then
            FinalRow = blckcll.Row - 1
            Exit For
        End If
    Next blckcll
    
    'delete all conditional formatting
    For Each del In ActiveSheet.Range(Cells(1, 1), Cells(FinalRow, 1))
        del.FormatConditions.Delete
    Next del
    
    Application.ScreenUpdating = True
    
    'create data bar conditional formatting
    Set CellDataBar = ActiveSheet.Range(Cells(3, 1), Cells(FinalRow, 1)).FormatConditions.AddDatabar
    Set CellDataBarColor = CellDataBar.BarColor
    
    'format data bars
    CellDataBarColor.Color = RGB(99, 195, 132)
    CellDataBar.BarFillType = xlDataBarFillSolid
    CellDataBar.MinPoint.Modify newtype:=xlConditionValuePercent, newvalue:=0
    CellDataBar.MaxPoint.Modify newtype:=xlConditionValuePercent, newvalue:=100
    
End Sub

Public Sub InitializeColoredCells()

    Dim cell As Range
    Dim page As Worksheet
    Dim blckcll As Range
    Dim FinalRow As Double
    
    'iterate all pages
    For Each page In Worksheets
        
        'if page is a Set page
        If Left(page.Name, 3) = "Set" Then
    
            'find final row
            For Each blckcll In ActiveSheet.Range("A:A")
                If blckcll.Interior.Color = RGB(9, 10, 11) Then
                    FinalRow = blckcll.Row - 1
                    Exit For
                End If
            Next blckcll
            
            'Column E
            For Each cell In ActiveSheet.Range(Cells(1, 5), Cells(FinalRow, 5))
                Application.GoTo cell
                If cell.Value = Worksheets("List Source").Cells(3, 1).Value Then
                    Call GreenCell
                ElseIf cell.Value = Worksheets("List Source").Cells(4, 1).Value Then
                    Call LightGreenCell
                ElseIf cell.Value = Worksheets("List Source").Cells(5, 1).Value Then
                    Call WhiteCell
                ElseIf cell.Value = Worksheets("List Source").Cells(6, 1).Value Then
                    Call OrangeCell
                ElseIf cell.Value = Worksheets("List Source").Cells(7, 1).Value Then
                    Call YellowCell
                ElseIf cell.Value = Worksheets("List Source").Cells(8, 1).Value Then
                    Call RedCell
                End If
            Next cell
            
            'Column F
            For Each cell In ActiveSheet.Range(Cells(1, 6), Cells(FinalRow, 6))
                Application.GoTo cell
                If cell.Value = Worksheets("List Source").Cells(3, 2).Value Then
                    Call WhiteCell
                ElseIf cell.Value = Worksheets("List Source").Cells(4, 2).Value Then
                    Call LightYellowCell
                ElseIf cell.Value = Worksheets("List Source").Cells(5, 2).Value Then
                    Call YellowCell
                ElseIf cell.Value = Worksheets("List Source").Cells(6, 2).Value Then
                    Call DarkYellowCell
                ElseIf cell.Value = Worksheets("List Source").Cells(7, 2).Value Then
                    Call BlueCell
                ElseIf cell.Value = Worksheets("List Source").Cells(8, 2).Value Then
                    Call RedCell
                End If
            Next cell
        
        End If
        
    Next page

End Sub

Public Sub GetRGBColor_Fill()

    Dim cel As Range
    Dim HEXcolor As String
    Dim RGBcolor As String

    For Each cel In Selection
        HEXcolor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)
        RGBcolor = "RGB (" & CInt("&H" & Right(HEXcolor, 2)) & ", " & CInt("&H" & Mid(HEXcolor, 3, 2)) & ", " & CInt("&H" & Left(HEXcolor, 2)) & ")"
        cel.Value = RGBcolor
    Next cel

End Sub


