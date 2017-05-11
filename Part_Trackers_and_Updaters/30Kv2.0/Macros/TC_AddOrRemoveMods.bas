Attribute VB_Name = "TC_AddOrRemoveMods"
Public Sub ExportMods()

    Dim objMyProj As Object
    Dim objVBComp As Object
    Dim i As Integer
     
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In Application.VBE.ActiveVBProject.VBComponents
        If (Left(objVBComp.Name, 3) = "TC_") Then
            'MsgBox objVBComp.Name
            objVBComp.Export "C:\Users\xrgb231\Documents\VBA - Excel Programming\Cody\NEW Tracker Update Program\Exported Macros\30K File\" & objVBComp.Name & ".bas"
        End If
    Next objVBComp

End Sub

Public Sub RemoveMods()

    Dim objMyProj As Object
    Dim objVBComp As Object
    Dim i As Integer
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In Application.VBE.ActiveVBProject.VBComponents
        If (Left(objVBComp.Name, 3) = "TC_") Then
            'MsgBox objVBComp.Name
            'Application.VBE.ActiveVBProject.VBComponents.Remove objVBComp  'UNCOMMENT LINE TO DELETE ALL MODS
        End If
    Next objVBComp

End Sub
