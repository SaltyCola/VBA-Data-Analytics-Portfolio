Attribute VB_Name = "SGV_CountSheet"
'Click Event Code is located in the "Sheet23 (Count Sheet)" Code

Public PartNumber As String 'Part Number to determine which tab to grab SN's from

Public Sub ClickButton080()

    PartNumber = "5319080"
    Call CountSheetAutoFill

End Sub

Public Sub ClickButton180()

    PartNumber = "5319180"
    Call CountSheetAutoFill

End Sub

Public Sub ClickButton280()

    PartNumber = "5319280"
    Call CountSheetAutoFill

End Sub

Public Sub ClickButton380()

    PartNumber = "5319380"
    Call CountSheetAutoFill

End Sub

Public Sub ClickButton480()

    PartNumber = "5319480"
    Call CountSheetAutoFill

End Sub

Public Sub ClearCountSheet()

    'unhide cells
    Application.ScreenUpdating = False
    Worksheets("Count Sheet").Rows("1:450").Hidden = False

    'reset Date, Part and Name # headers
    Worksheets("Count Sheet").Cells(1, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(1, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(1, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(46, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(46, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(46, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(91, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(91, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(91, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(136, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(136, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(136, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(181, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(181, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(181, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(226, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(226, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(226, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(271, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(271, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(271, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(316, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(316, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(316, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(361, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(361, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(361, 7).Value = "Name:" & Space(63)
    Worksheets("Count Sheet").Cells(406, 1).Value = "Date: "
    Worksheets("Count Sheet").Cells(406, 3).Value = "Part #: "
    Worksheets("Count Sheet").Cells(406, 7).Value = "Name:" & Space(63)

    'clear all table values
    Range("A3:H45").ClearContents
    Range("A48:H90").ClearContents
    Range("A93:H135").ClearContents
    Range("A138:H180").ClearContents
    Range("A183:H225").ClearContents
    Range("A228:H270").ClearContents
    Range("A273:H315").ClearContents
    Range("A318:H360").ClearContents
    Range("A363:H405").ClearContents
    Range("A408:H450").ClearContents
    
    'turn on screen updating
    Application.ScreenUpdating = True

End Sub

Public Sub CountSheetAutoFill()

    'declare variables
    Dim todaysDate As String 'String of today's date to input into the header of each sheet
    Dim snList As Collection 'SGV_SN object collection for printing to count sheet
    Dim c As Range 'Generic range iteration object for "for each" loops
    Dim d As Range 'Generic range iteration object for 2nd tier "for each" loops
    Dim e As Range 'Generic range iteration object for 3rd tier "for each" loops
    Dim snRow As Double 'current wrk's "S/N" row
    Dim OpRow As Double 'current operation row being searched
    Dim sOpRow As Double 'starting op row for iteration
    Dim eOpRow As Double 'ending op row for iteration

    'call count sheet clear sub
    Call ClearCountSheet
    
    'disable screen updating
    Application.ScreenUpdating = False
    
    'activate correct worksheet
    Worksheets(PartNumber).Activate
    
    'find S/N row, sOpRow and eOpRow
    For Each c In Worksheets(PartNumber).Range(Cells(10, 2), Cells(40, 2))
        'S/N row found
        If c.Value = "S/N" Then
            snRow = c.Row
        End If
        'sOpRow found
        If c.Value = "Shipped" Then
            sOpRow = c.Row
        End If
        'eOpRow found
        If c.Value = "Launch" Then
            eOpRow = c.Row
        End If
    Next c
    
    'initialize SGV_SN Collection
    Set snList = New Collection
    
    'declare SGV_SN transfer object
    Dim transSN As SGV_SN
    
    
    
    'iterate S/N row to add SGV_SN object to array
    For Each c In Worksheets(PartNumber).Range(snRow & ":" & snRow)
        
        'make transSN nothing
        Set transSN = Nothing
        
        'exit for if SN is empty
        If (c.Column > 2) And (c.EntireColumn.Hidden = False) And (IsEmpty(c)) Then
            Exit For
        
        'get SGV_SN object properties and add to collection
        ElseIf (c.Column > 2) And (c.EntireColumn.Hidden = False) And Not (IsEmpty(c)) Then
            
            'create transSN object
            Set transSN = New SGV_SN
            transSN.SerialNumber = Right(c.Value, 5)
            
            For Each d In Range(Cells(sOpRow, c.Column), Cells(eOpRow, c.Column))
                
                'yellow found
                If (Cells(d.Row, 2).Value = Worksheets("Count Sheet").Range("O12").Value) And (d.Interior.Color = Worksheets("Count Sheet").Range("O12").Interior.Color) And (d.Offset(-1, 0).Interior.Color = RGB(255, 255, 255)) Then
                    transSN.LastOp = Worksheets("Count Sheet").Range("Q12").Value
                    transSN.LastDate = Str(d.Value)
                    Exit For
                
                'green found
                ElseIf ((d.Interior.Color = RGB(146, 208, 80)) And (d.Row = sOpRow)) Or ((d.Interior.Color = RGB(146, 208, 80)) And (d.Offset(-1, 0).Interior.Color = RGB(255, 255, 255))) Then
                    transSN.LastDate = Str(d.Value)
                    For Each e In Worksheets("Count Sheet").Range("O3:O19")
                        'found op name
                        If e.Value = Worksheets(PartNumber).Cells(d.Row, 2).Value Then
                            transSN.LastOp = e.Offset(0, 2).Value
                            Exit For
                        End If
                    Next e
                    Exit For
                End If
                
            Next d
            
            'add SGV_SN object to collection
            snList.Add transSN
        
        End If
    Next c
    
    'activate count sheet worksheet
    Worksheets("Count Sheet").Activate
    
    'turn on screen updating
    Application.ScreenUpdating = True

    'initialize variables
    todaysDate = Str(Date)

    'apply todaysDate, PartNumber and Username to count sheets
    Worksheets("Count Sheet").Cells(1, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(1, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(1, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(46, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(46, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(46, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(91, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(91, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(91, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(136, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(136, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(136, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(181, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(181, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(181, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(226, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(226, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(226, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(271, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(271, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(271, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(316, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(316, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(316, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(361, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(361, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(361, 7).Value = "Name:   " & Application.UserName
    Worksheets("Count Sheet").Cells(406, 1).Value = "Date: " & todaysDate
    Worksheets("Count Sheet").Cells(406, 3).Value = "Part #: " & PartNumber
    Worksheets("Count Sheet").Cells(406, 7).Value = "Name:   " & Application.UserName
    
    'print out SN info
    i = 1
    For Each c In Worksheets("Count Sheet").Range("A:A")
        If i > snList.Count Then
            Exit For
        ElseIf (IsEmpty(c)) And Not (snList(i).SerialNumber = "") Then
            c.Value = snList(i).SerialNumber
            c.Offset(0, 1).Value = snList(i).LastOp
            c.Offset(0, 2).Value = snList(i).LastDate
            i = i + 1
        End If
    Next c

    'hide unused pages
    If Worksheets("Count Sheet").Range("A3").Value = "" Then: Worksheets("Count Sheet").Rows("1:45").Hidden = True
    If Worksheets("Count Sheet").Range("A48").Value = "" Then: Worksheets("Count Sheet").Rows("46:90").Hidden = True
    If Worksheets("Count Sheet").Range("A93").Value = "" Then: Worksheets("Count Sheet").Rows("91:135").Hidden = True
    If Worksheets("Count Sheet").Range("A138").Value = "" Then: Worksheets("Count Sheet").Rows("136:180").Hidden = True
    If Worksheets("Count Sheet").Range("A183").Value = "" Then: Worksheets("Count Sheet").Rows("181:225").Hidden = True
    If Worksheets("Count Sheet").Range("A228").Value = "" Then: Worksheets("Count Sheet").Rows("226:270").Hidden = True
    If Worksheets("Count Sheet").Range("A273").Value = "" Then: Worksheets("Count Sheet").Rows("271:315").Hidden = True
    If Worksheets("Count Sheet").Range("A318").Value = "" Then: Worksheets("Count Sheet").Rows("316:360").Hidden = True
    If Worksheets("Count Sheet").Range("A363").Value = "" Then: Worksheets("Count Sheet").Rows("361:405").Hidden = True
    If Worksheets("Count Sheet").Range("A408").Value = "" Then: Worksheets("Count Sheet").Rows("406:450").Hidden = True

End Sub


