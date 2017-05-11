Attribute VB_Name = "Mod_SaveAndZip"
Option Explicit

'================My Public Variables================'
Public dirAsBuiltFiles As String
Public nameTodaysFolder As String
Public dirTodaysFolder As String
Public iterGroup As Integer
Public nameGroupFolder As String
Public dirGroupFolder As String
Public cntGroupFile As Integer
Public nameGroupFile As String
Public nameSNFile As String
Public strPrtNum As String
Public strSerNum As String
'================My Public Variables================'

'TO CHANGE MASTER FILE SEARCH: 30K As Built

'Locate or Create directories for 10 SN's apiece
Sub CreateDirectories()
Attribute CreateDirectories.VB_ProcData.VB_Invoke_Func = "q\n14"

    'define main directory variables
    dirAsBuiltFiles = "U:\5. Cell Data\PWAA_GTF\Shipping\PW1100G As Built" '30K As Built Main Directory
    nameTodaysFolder = Replace(Str(Date), "/", ".")
    dirTodaysFolder = dirAsBuiltFiles & "\" & nameTodaysFolder
    
    'create todays date directory if it doesn't exist already
    If Len(Dir(dirTodaysFolder, vbDirectory)) = 0 Then
        MkDir dirTodaysFolder
    End If
    
    'Define Serial Number Information
    strPrtNum = ActiveWorkbook.Worksheets("As Built Data Form").Range("A2").Value
    strSerNum = ActiveWorkbook.Worksheets("As Built Data Form").Range("B2").Value
    
    'find current group folder
    For iterGroup = 1 To 100
        nameGroupFolder = "Group" & iterGroup
        dirGroupFolder = dirTodaysFolder & "\" & nameGroupFolder
        'create new group folder
        If Len(Dir(dirGroupFolder, vbDirectory)) = 0 Then
            MkDir dirGroupFolder
        End If
        'count files in group folder
        cntGroupFile = 0
        nameGroupFile = Dir(dirGroupFolder & "\*.xls")
        Do While nameGroupFile <> ""
            cntGroupFile = cntGroupFile + 1
            nameGroupFile = Dir()
        Loop
        'save file to group folder up to 10
        If cntGroupFile < 10 Then
            Call SaveFileName
            Exit For
        End If
    Next iterGroup
    
End Sub

'Save File to current group directory until there are 10 files
Sub SaveFileName()

    'Create new template file
    Workbooks.Open ("U:\5. Cell Data\PWAA_GTF\FBC_Operations\Production Tracking\As Built Data" & "\" & "As Built Template.xlsx")
    
    'delete template's As Built Data Form tab
    Application.DisplayAlerts = False
    Workbooks("As Built Template").Worksheets("As Built Data Form").Delete
    Application.DisplayAlerts = True
    
    'copy over SN Info
    Workbooks("30K Data_Master_RGBSI").Worksheets("As Built Data Form").Copy Workbooks("As Built Template").Worksheets("As Built Data Requirements") '30K As Built Master File copying to template
    
    'delete formulas
    Dim cll As Range
    For Each cll In Workbooks("As Built Template").Worksheets("As Built Data Form").Range("A1:I6")
        cll.Value = cll.Value2
    Next cll
    
    'save file in group location
    ActiveWorkbook.SaveAs Filename:=dirGroupFolder & "\" & strPrtNum & "_" & strSerNum
    
    'close file
    Workbooks(strPrtNum & "_" & strSerNum).Close
    
    'file saved message
    MsgBox "Saved File: '" & strPrtNum & "_" & strSerNum & "'."

End Sub

