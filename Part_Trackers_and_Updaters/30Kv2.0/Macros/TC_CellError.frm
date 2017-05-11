VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TC_CellError 
   Caption         =   "Cell Error Found!"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2745
   OleObjectBlob   =   "TC_CellError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TC_CellError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rngCellAddress As Range 'Holds cell range where error occurred.
Private bRestrictedToDateType As Boolean 'True: Text in txtCellValue must be a specific value type.
Private strAllowedDateChars As String 'Holds string of the characters allowed for a date type.
Private bContinue As Boolean 'True: Allow userform to hide, clear, and continue reading subroutine.

Private Sub btnContinue_Click()
'Validate text for applying to cell, apply to cell, continue with reading subroutine.

    'Validate cell value textbox
    Me.Validate_txtCellValue
    
    'Continue?
    If bContinue Then
    
        'Apply cell value textbox text to cell range
        rngCellAddress.Value = Me.txtCellValue.Text
        
        'Hide and clear userform, then continue reading subroutine
        Me.Hide
        Me.ClearCellErrorData
    
    End If

End Sub

Public Sub GrabCellValueError(ByVal CellError_Row As Integer, ByVal CellError_Column As Integer, ByVal CellError_RestrictToDateType As Boolean)
'To be called before showing the userform, loads values into userform.

    'set private restriction variables
    Set rngCellAddress = Cells(CellError_Row, CellError_Column)
    bRestrictedToDateType = CellError_RestrictToDateType
    
    'Go to cell
    Application.Goto rngCellAddress, Scroll:=True
    
    'grab cell data
    Me.txtCellAddress = CStr(rngCellAddress.Address)
    Me.txtCellValue = CStr(rngCellAddress.Value)

End Sub

Public Sub Validate_txtCellValue()
'Validate text before applying to cell.


    Dim i As Integer 'iterator
    
    
    'initialize continue boolean
    bContinue = True
    
    
    'only need to validate if the cell's value is restricted to a date type
    If bRestrictedToDateType Then
    
        'iterate txtCellValue.Text
        For i = 1 To Len(Me.txtCellValue.Text)
            If InStr(1, strAllowedDateChars, Mid(Me.txtCellValue.Text, i, 1)) = 0 Then
                'Error Message
                MsgBox "The cell value textbox has an invalid character: '" & Mid(Me.txtCellValue.Text, i, 1) & "'", , "Data Validation Error!"
                'prevent continue
                bContinue = False
                Exit Sub
            End If
        Next i
        
        'Check if valid date
        If Me.txtCellValue.Text <> "" And Not IsDate(Me.txtCellValue.Text) Then
            'Error Message
            MsgBox "The cell value textbox is not a date.", , "Data Validation Error!"
            'prevent continue
            bContinue = False
            Exit Sub
        End If
    
    End If


End Sub

Public Sub ClearCellErrorData()
'Clears out Userform for future use.

    'Clear textboxes
    Me.txtCellAddress.Text = ""
    Me.txtCellValue.Text = ""
    
    'Reset private variables
    Set rngCellAddress = Nothing
    bRestrictedToDateType = False
    strAllowedDateChars = "/0123456789"
    bContinue = False

End Sub

Private Sub UserForm_Initialize()
'initialize userform for use

    'initialize private variables
    Set rngCellAddress = Nothing
    bRestrictedToDateType = False
    strAllowedDateChars = "/0123456789"
    bContinue = False

End Sub

Private Sub UserForm_Activate()
'Highlight text in the cell value textbox to show if there is a space.

    Me.txtCellValue.SetFocus
    Me.txtCellValue.SelStart = 0
    Me.txtCellValue.SelLength = Len(Me.txtCellValue.Text)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent userform close on red x click.

    If CloseMode = 0 Then: Cancel = True

End Sub
