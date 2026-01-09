Attribute VB_Name = "ExportToGit"
Option Explicit

Sub ExportVbaForGit()
    Dim cmp As VBIDE.VBComponent
    Dim vbaFolderPath As String
    Dim targetPath As String
    Dim outExt As String
    Dim fileExport As String
    
    ' koøen GIT složky
    vbaFolderPath = "C:\GitHub\ineko\VBA\Rozpocet\"
    
    ' vytvoø podsložky, pokud neexistují
    CreateFolderIfMissing vbaFolderPath
    CreateFolderIfMissing vbaFolderPath & "ExcelObjects\"
    CreateFolderIfMissing vbaFolderPath & "Forms\"
    CreateFolderIfMissing vbaFolderPath & "Modules\"
    
    On Error GoTo MustTrustVBAProject
    Set cmp = ThisWorkbook.VBProject.VBComponents(1)
    On Error GoTo 0
    
    For Each cmp In ThisWorkbook.VBProject.VBComponents
        
        Select Case cmp.Type
            Case vbext_ct_Document
                ' listy + ThisWorkbook
                targetPath = vbaFolderPath & "ExcelObjects\"
                outExt = ".cls"
                
            Case vbext_ct_MSForm
                ' UserFormy
                targetPath = vbaFolderPath & "Forms\"
                outExt = ".frm"
                
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' standard + class moduly
                targetPath = vbaFolderPath & "Modules\"
                outExt = IIf(cmp.Type = vbext_ct_StdModule, ".bas", ".cls")
                
            Case Else
                targetPath = ""
                outExt = ""
        End Select
        
        If targetPath <> "" And outExt <> "" Then
            fileExport = targetPath & cmp.Name & outExt
            If Dir(fileExport) <> "" Then Kill fileExport
            cmp.Export fileExport
        End If
    Next cmp
    
    MsgBox "Export hotový do: " & vbaFolderPath, vbInformation
    Exit Sub

MustTrustVBAProject:
    MsgBox "Zapni prosím 'Trust access to the VBA project object model' v Trust Center.", _
           vbCritical, "ExportVbaForGit"
End Sub

Private Sub CreateFolderIfMissing(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub


