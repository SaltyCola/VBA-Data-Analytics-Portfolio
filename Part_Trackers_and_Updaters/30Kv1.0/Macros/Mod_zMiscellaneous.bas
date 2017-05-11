Attribute VB_Name = "Mod_zMiscellaneous"
Option Explicit

Sub FixDateCellYears()
Attribute FixDateCellYears.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim cell As Range
    
    For Each cell In ActiveSheet.Range("C7:SL43")
        Application.Goto cell, Scroll:=True
        If Not IsEmpty(cell) Then
            
            If cell.Text = "#REF!" Or cell.Text = "#VALUE!" Or cell.Text = "######" Then
                cell.Value = ""
                MsgBox cell.Value
            End If
            
            If Right(cell.Value, 13) = "  12:00:00 PM" Then
                cell.Value = Left(cell.Value, (Len(cell.Value) - 13))
                MsgBox cell.Value
            ElseIf Right(cell.Value, 12) = " 12:00:00 PM" Then
                cell.Value = Left(cell.Value, (Len(cell.Value) - 12))
                MsgBox cell.Value
            End If
            
            If Not IsEmpty(cell) And Not ((Right(cell.Value, 4) = "2016") Or (Right(cell.Value, 4) = "2015")) Then
                cell.Value = Left(cell.Value, (Len(cell.Value) - 4)) & 2015
                MsgBox cell.Value
            End If
            
        End If
    Next cell

End Sub

Sub EraseFormulasAndBlankOutColor()

    Dim c As Double
    Dim r As Double
    Dim cll As Range
    Dim ldTime As Double
    
    For c = 520 To 3 Step -1
        For r = 43 To 7 Step -1
            Set cll = ActiveSheet.Range(Cells(r, c), Cells(r, c))
            If Not (cll.Interior.Color = RGB(146, 208, 80)) And Not (cll.Interior.Color = RGB(79, 98, 40)) And Not (cll.Interior.Color = RGB(196, 215, 155)) And Not (cll.Interior.Color = RGB(0, 176, 80)) And Not (cll.Interior.Color = RGB(255, 192, 0)) And Not (cll.Interior.Color = RGB(146, 205, 220)) And Not (cll.Interior.Color = RGB(255, 0, 0)) And Not (cll.Interior.Color = RGB(0, 0, 0)) Then
                Application.Goto cll
                cll.Interior.Color = RGB(255, 255, 255)
                'set lead time variable
                If Worksheets("NEO 5322121").Cells(r, 1).Value = 0.5 Then
                    ldTime = 0
                Else
                    ldTime = Worksheets("NEO 5322121").Cells(r, 1).Value
                End If
                'todays date for bottom row or row above an empty colored cell
                If r = 43 Or cll.Offset(1, 0).Value = "" Then
                    cll.Value = Date
                Else
                    cll.Value = cll.Offset(1, 0).Value + ldTime
                End If
            End If
        Next r
    Next c

End Sub

Sub NewFormDeclarationTest()

    Dim frmTEST As MsgWaterfall
    Dim i As Integer
    
    For i = 1 To 10
        Set frmTEST = New MsgWaterfall
        frmTEST.Show vbModeless
        DoEvents
        frmTEST.Hide
    Next i

End Sub

Sub GetRGBColor_Fill()

    Dim cel As Range
    Dim HEXcolor As String
    Dim RGBcolor As String

    For Each cel In Selection
        HEXcolor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)
        RGBcolor = "RGB (" & CInt("&H" & Right(HEXcolor, 2)) & ", " & CInt("&H" & Mid(HEXcolor, 3, 2)) & ", " & CInt("&H" & Left(HEXcolor, 2)) & ")"
        cel.Value = RGBcolor
    Next cel

End Sub

Sub DateAddingTest()

    MsgBox ActiveCell.Value
    MsgBox ActiveCell.Value - 0.5
    ActiveCell.Offset(1, 0).Value = ActiveCell.Value - 0.5

End Sub

Sub ValueTest()
    
    MsgBox ActiveCell.Value
    MsgBox ActiveCell.Value2
    MsgBox ActiveCell.Text

End Sub

Sub CutInsertTest()

    Worksheets("SN BACKUP SHEET").Columns(3).Cut
    Worksheets("SN BACKUP SHEET").Columns(2).Insert
    Worksheets("SN BACKUP SHEET").Columns(2).Insert
    Worksheets("SN BACKUP SHEET").Columns(2).Interior.Color = RGB(0, 0, 0)
    MsgBox ("Test Delete Function")
    Worksheets("SN BACKUP SHEET").Columns(2).Delete

End Sub

Sub rangetest()

    Dim rng As Range
    Dim i As Range
    
    Set rng = Worksheets("SN BACKUP SHEET").Rows(10)
    
    For Each i In rng
        i.Value = "YES"
    Next i

End Sub

Sub FixAllLastYearsDates()

    Dim cll As Range
    Dim rngSheet As Range
    Dim LastYearDateIncorrect As String
    Dim LastYearDateFixed As String
    Dim errortxt As String
    Dim newtxt As String
    
    Set rngSheet = Worksheets("NEO 5322121").Range("C7", "RT43")

    For Each cll In rngSheet
        'If (cll.Row <= 33) Or (cll.Row >= 38 And cll.Row <= 43) Then
            If cll.Text = "#REF!" Then
                MsgBox (cll.Address)
                errortxt = cll.Text
                newtxt = Str(cll.Offset(0, -1))
                cll.Value = cll.Offset(0, -1)
                MsgBox (errortxt & " has been changed to: " & newtxt)
            End If
            If Not (Left(Str(cll.Value), 2) = "1/") And Not (Left(Str(cll.Value), 2) = "2/") And Not (Left(Str(cll.Value), 2) = "3/") And Not (Left(Str(cll.Value), 2) = "4/") And Not (Left(Str(cll.Value), 2) = "5/") Then
                On Error GoTo lineErrorAddress
                If Right(Str(cll.Value), 5) = "/2016" Then
                    LastYearDateIncorrect = Str(cll.Value)
                    LastYearDateFixed = Left(Str(cll.Value), Len(Str(cll.Value)) - 5) & "/2015"
                    cll.Value = LastYearDateFixed
                    MsgBox (LastYearDateIncorrect & " has been changed to: " & LastYearDateFixed)
                End If
                If False Then
lineErrorAddress:
                    MsgBox (cll.Address)
                    errortxt = cll.Text
                    newtxt = Str(cll.Offset(0, -1))
                    cll.Value = cll.Offset(0, -1)
                    MsgBox (errortxt & " has been changed to: " & newtxt)
                End If
            End If
        'End If
    Next cll
    
End Sub

Sub CreateBackupDirectoryTest()

    Dim MyPath As String
    Dim MyDate
    Dim MyTime
    Dim dateStr As String
    Dim timeStr As String
    Dim MyBackupPath As String
    
    'define variables
    MyPath = ActiveWorkbook.Path
    MyDate = Date
    MyTime = Time
    dateStr = Format(MyDate, "DD-MM-YYYY")
    timeStr = Format(MyTime, "hh.mm.ss")
    
    'create backup folder if DNE
    If Len(Dir(MyPath & "\BACKUPS - 30K Update Program", vbDirectory)) = 0 Then
        MkDir (MyPath & "\BACKUPS - 30K Update Program")
        MsgBox ("Directory Created!")
    Else
        MsgBox ("Directory Already Exists!")
    End If

End Sub

Sub FixTextFormatToDate()

    Dim rngFix As Range
    Dim fx As Range
    Set rngFix = Worksheets("NEO 5322121").Range("C52:PD52")

    For Each fx In rngFix
        
        If Len(fx.Value) < 5 Or Not Right(fx.Value, 5) = "/2016" Then
            fx.Value = fx.Value & "/2016"
        End If
        
    Next fx

End Sub

Public Sub TestRangeColor()

    If Worksheets("NEO 5322121").Cells(11, 20).Interior.Color = RGB(255, 255, 255) Then
        MsgBox "Yep"
    End If

End Sub

Sub TestAddressFunction()

    MsgBox ActiveCell.EntireColumn.Address

End Sub

Sub ExportMods()

    Dim objMyProj As Object
    Dim objVBComp As Object
    Dim i As Integer
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In Application.VBE.ActiveVBProject.VBComponents
        If (Left(objVBComp.Name, 3) = "Mod") Or (Left(objVBComp.Name, 3) = "SN_") Or (Left(objVBComp.Name, 3) = "Msg") Then
            'MsgBox objVBComp.Name
            objVBComp.Export "C:\Users\xrgb231\Documents\VBA - Excel Programming\Andrea\30K Tracker Update Program\Exported Macro Files\" & objVBComp.Name & ".bas"
        End If
    Next objVBComp

End Sub

Sub RemoveMods()

    Dim objMyProj As Object
    Dim objVBComp As Object
    Dim i As Integer
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In Application.VBE.ActiveVBProject.VBComponents
        If (Left(objVBComp.Name, 3) = "Mod") Or (Left(objVBComp.Name, 3) = "SN_") Or (Left(objVBComp.Name, 3) = "Msg") Then
            'MsgBox objVBComp.Name
            'Application.VBE.ActiveVBProject.VBComponents.Remove objVBComp  'UNCOMMENT LINE TO DELETE ALL MODS
        End If
    Next objVBComp

End Sub
