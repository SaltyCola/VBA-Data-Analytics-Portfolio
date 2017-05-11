VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    'if cell in header row being changed
    If Not (Intersect(Target, Range("8:8")) Is Nothing) And (Intersect(Target, Range("9:9")) Is Nothing) And (Intersect(Target, Range("7:7")) Is Nothing) Then
        Select Case Target
            Case ActiveSheet.Range("A43").Value: Call pShipped(Target)
            Case ActiveSheet.Range("A44").Value: Call oCustomerHold(Target)
            Case ActiveSheet.Range("A45").Value: Call gWIP(Target)
        End Select
    End If

End Sub

Private Sub pShipped(ByVal Target As Range)

    Dim c As Range 'generic iteration range object

    Target.Interior.Color = ActiveSheet.Range("A43").Interior.Color
    Target.Offset(1, 0).Interior.Color = ActiveSheet.Range("A43").Interior.Color
    Target.Offset(2, 0).Interior.Color = ActiveSheet.Range("A43").Interior.Color
    Target.Offset(3, 0).Interior.Color = ActiveSheet.Range("A43").Interior.Color
    Target.Offset(5, 0).Interior.Color = ActiveSheet.Range("A43").Interior.Color
    
    'if sales order is in "On Hold List" tab, white out row
    For Each c In Worksheets("On Hold List").Range("A6:A500")
        If IsEmpty(c) Then
            Exit For
        ElseIf c.Value = Target.Offset(5, 0).Value Then
            c.Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 1).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 2).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 3).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 4).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 5).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 6).Interior.Color = RGB(255, 255, 255)
        End If
    Next c

End Sub

Private Sub oCustomerHold(ByVal Target As Range)

    Dim c As Range 'generic iteration range object

    Target.Interior.Color = ActiveSheet.Range("A44").Interior.Color
    Target.Offset(1, 0).Interior.Color = ActiveSheet.Range("A44").Interior.Color
    Target.Offset(2, 0).Interior.Color = ActiveSheet.Range("A44").Interior.Color
    Target.Offset(3, 0).Interior.Color = ActiveSheet.Range("A44").Interior.Color
    Target.Offset(5, 0).Interior.Color = RGB(255, 255, 255)
    
    'if sales order is in "On Hold List" tab, white out row
    For Each c In Worksheets("On Hold List").Range("A6:A500")
        If IsEmpty(c) Then
            Exit For
        ElseIf c.Value = Target.Offset(5, 0).Value Then
            c.Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 1).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 2).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 3).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 4).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 5).Interior.Color = RGB(255, 255, 255)
            c.Offset(0, 6).Interior.Color = RGB(255, 255, 255)
        End If
    Next c

End Sub

Private Sub gWIP(ByVal Target As Range)

    Dim c As Range 'generic iteration range object
    
    Target.Interior.Color = ActiveSheet.Range("A45").Interior.Color
    Target.Offset(1, 0).Interior.Color = RGB(255, 255, 255)
    Target.Offset(2, 0).Interior.Color = RGB(255, 255, 255)
    Target.Offset(3, 0).Interior.Color = RGB(255, 255, 255)
    Target.Offset(5, 0).Interior.Color = RGB(255, 255, 255)
    
    'if sales order is in "On Hold List" tab, green out row
    For Each c In Worksheets("On Hold List").Range("A6:A500")
        If IsEmpty(c) Then
            Exit For
        ElseIf c.Value = Target.Offset(5, 0).Value Then
            c.Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 1).Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 2).Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 3).Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 4).Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 5).Interior.Color = RGB(146, 208, 80)
            c.Offset(0, 6).Interior.Color = RGB(146, 208, 80)
        End If
    Next c

End Sub
