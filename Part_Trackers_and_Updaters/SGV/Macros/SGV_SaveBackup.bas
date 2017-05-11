Attribute VB_Name = "SGV_SaveBackup"
Public Sub SaveFileBackup()
'saves a backup file in backup directory in personal computer documents

    Dim frmMsgSavingBackupFile As MsgSavingBackupFile
    Dim MyPath As String
    Dim MyNow As String
    Dim nowStr As String
    Dim MyBackupPath As String
    
    'define variables
    MyPath = Environ$("USERPROFILE") & "\Documents" 'grab local documents folder path!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    MyNow = Str(Now)
    nowStr = Replace(Replace(MyNow, ":", "."), "/", "-")
    
    'show Saving Backup Message
    Set frmMsgSavingBackupFile = New MsgSavingBackupFile
    frmMsgSavingBackupFile.Show vbModeless
    DoEvents
    
    'deactivate alerts
    Application.DisplayAlerts = False
    
'    'delete previous backup folders
'    Dim i As Double
'    Dim Fs As Object
'    For i = 1 To 7
'        If Not Len(Dir(MyPath & "\BACKUPS - SGV Tracker " & Replace(Str(DateAdd("d", -(i), Date)), "/", "-"), vbDirectory)) = 0 Then
'            Set Fs = CreateObject("Scripting.FileSystemObject")
'            Fs.DeleteFolder (MyPath & "\BACKUPS - SGV Tracker " & Replace(Str(DateAdd("d", -(i), Date)), "/", "-")), True
'        End If
'    Next i
    
    'create backup folder if DNE
    If Len(Dir(MyPath & "\BACKUPS - SGV Tracker " & Replace(Str(Date), "/", "-"), vbDirectory)) = 0 Then
        MkDir (MyPath & "\BACKUPS - SGV Tracker " & Replace(Str(Date), "/", "-"))
    End If
    
    'define backup path
    MyBackupPath = (MyPath & "\BACKUPS - SGV Tracker " & Replace(Str(Date), "/", "-") & "\")
    
    'save backup copy
    ActiveWorkbook.SaveCopyAs Filename:=MyBackupPath & " (" & nowStr & ") " & ActiveWorkbook.Name
    
    'save current file
    ActiveWorkbook.Save
    
    'reactivate alerts
    Application.DisplayAlerts = True
    
    'close Saving Backup Message
    frmMsgSavingBackupFile.Hide
    Set frmMsgSavingBackupFile = Nothing
    

End Sub
