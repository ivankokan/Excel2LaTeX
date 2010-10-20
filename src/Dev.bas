Attribute VB_Name = "Dev"
Option Explicit

Public Sub Diff()
    PrepareCommit
    VBA.Shell Printf("bzr qdiff ""%1..""", BaseDir()), vbNormalFocus
End Sub

Public Sub Commit(Optional ByVal sMessage As String)
    PrepareCommit
    If sMessage <> "" Then
        VBA.Shell Printf("bzr ci ""%1.."" -m ""%2""", BaseDir(), Replace(sMessage, """", "'"))
    Else
        VBA.Shell Printf("bzr ci ""%1..""", BaseDir())
    End If
End Sub

Public Sub CreateBzrWorksheet()
    Dim sTargetPath As String
    sTargetPath = BaseDir() & "Excel2LaTeX.xls"
    
    Application.ActiveWorkbook.SaveCopyAs sTargetPath
    
    Dim pTargetWkBook As Workbook
    Set pTargetWkBook = Application.Workbooks.Open(sTargetPath)
    
    DropCodeModules pTargetWkBook
    ImportCodeModules pTargetWkBook, True, True
    
    pTargetWkBook.Close True
End Sub

Public Sub CreateDevWorksheet()
    Dim sTargetPath As String
    sTargetPath = BaseDir() & "Excel2LaTeXDev.xls"
    
    Application.ActiveWorkbook.SaveCopyAs sTargetPath
    
    Dim pOldWkBook As Workbook
    Set pOldWkBook = Application.ActiveWorkbook
    
    Dim pTargetWkBook As Workbook
    Set pTargetWkBook = Application.Workbooks.Open(sTargetPath)
    
    DropCodeModules pTargetWkBook
    ImportCodeModules pTargetWkBook, True
    
    pTargetWkBook.Save
    
    pOldWkBook.Close False
End Sub

Private Sub PrepareCommit()
    ActiveWorkbook.Save
    ExportToCodeModules
End Sub

Public Sub ExportToAddin()
    Const TEMPLATE_FILE = "Template.xla.xls"
    Const ADDIN_PATH As String = "..\Excel2LaTeX.xla"
    
    ExportToNewSheet TEMPLATE_FILE, ADDIN_PATH
End Sub
    
Private Sub ExportToNewSheet(ByVal sTemplateFile As String, ByVal sTargetFileName As String)
    Dim sDir As String
    sDir = BaseDir()
    
    Dim sTargetPath As String
    sTargetPath = sDir & sTargetFileName
    
    VBA.FileSystem.FileCopy sDir & sTemplateFile, sTargetPath
    
    Dim pTargetWkBook As Workbook
    Set pTargetWkBook = Application.Workbooks.Open(sTargetPath)
    
    ImportCodeModules pTargetWkBook, False
    
    pTargetWkBook.Close True
End Sub
    
Private Sub ImportCodeModules(ByVal pTargetWkBook As Workbook, ByVal bImportDevModule As Boolean, Optional ByVal bImportOnlyDevModule As Boolean = False)
    Dim sDir As String
    sDir = BaseDir()
    
    Dim sCurrentFileName As String
    sCurrentFileName = VBA.FileSystem.Dir(sDir, vbNormal)
    Do While sCurrentFileName <> ""
        If (Not bImportDevModule) And (sCurrentFileName = "Dev.bas") Then
            ' Ignore development module
        ElseIf bImportOnlyDevModule And sCurrentFileName <> "Dev.bas" Then
            ' Ignore anything but development module
        ElseIf sCurrentFileName Like "*.bas" Or sCurrentFileName Like "*.frm" Or sCurrentFileName Like "*.cls" Then
            ImportComponent pTargetWkBook, sDir, sCurrentFileName
        End If
        
        sCurrentFileName = VBA.FileSystem.Dir()
    Loop
End Sub

Private Sub DropCodeModules(ByVal pTargetWkBook As Workbook)
    Dim pVbComponent As VBComponent
    For Each pVbComponent In pTargetWkBook.VBProject.VBComponents
        Select Case GetFileExtension(pVbComponent)
        Case ".bas", ".cls", ".frm"
            pTargetWkBook.VBProject.VBComponents.Remove pVbComponent
        End Select
    Next
End Sub

Private Sub ExportToCodeModules()
    Dim sDir As String
    sDir = BaseDir()
    
    Dim pVbComponent As VBComponent
    For Each pVbComponent In Application.VBE.ActiveVBProject.VBComponents
        ExportComponent sDir, pVbComponent
    Next
End Sub





Private Sub ImportComponent(ByVal pTargetWkBook As Workbook, ByVal sDir As String, ByVal sFileName As String)
    pTargetWkBook.VBProject.VBComponents.Import sDir & sFileName
End Sub



Private Sub ExportComponent(ByVal sDir As String, ByVal pVbComponent As VBComponent)
    Dim sName As String
    sName = pVbComponent.Name
    
    Dim sOldName As String
    sOldName = sName & ".old"
    
    Dim sExtension As String
    sExtension = GetFileExtension(pVbComponent)
    
    Dim sDualExtension As String
    
    Select Case sExtension
    Case ".frm"
        ' The .frx file changes with every export of the module.
        ' To prevent this, we check if the code for the form has changed.
        ' If not, we assume that the form hasn't changed at all,
        ' and revert to the previous state.
        ' CAVE: This means that, e.g., after changing only the position of a control,
        ' the .frx file WILL NOT be updated. If you edit the form,
        ' always make at least a no-op change to the code module, e.g., add a line at the end.
        sDualExtension = ".frx"
        
        VBA.FileSystem.FileCopy sDir & sName & sExtension, sDir & sOldName & sExtension
        VBA.FileSystem.FileCopy sDir & sName & sDualExtension, sDir & sOldName & sDualExtension
        
        pVbComponent.Export sDir & sName & sExtension
        
        If FilesEqual(sDir & sName & sExtension, sDir & sOldName & sExtension) Then
            VBA.FileSystem.FileCopy sDir & sOldName & sExtension, sDir & sName & sExtension
            VBA.FileSystem.FileCopy sDir & sOldName & sDualExtension, sDir & sName & sDualExtension
        End If
        
        VBA.FileSystem.Kill sDir & sOldName & sExtension
        VBA.FileSystem.Kill sDir & sOldName & sDualExtension
    Case ""
        ' Skip this kind of module
    Case Else
        pVbComponent.Export sDir & sName & sExtension
    End Select
End Sub






Private Function BaseDir() As String
    SplitPath Application.VBE.ActiveVBProject.FileName, sDir:=BaseDir
End Function


Private Function FilesEqual(ByVal sFile1 As String, ByVal sFile2 As String) As Boolean
    If SafeFileLen(sFile1) <> SafeFileLen(sFile2) Then Exit Function
    
    FilesEqual = True
    
    Dim lFile1 As Long
    lFile1 = FreeFile
    Open sFile1 For Input Access Read Lock Write As #lFile1
    Dim lFile2 As Long
    lFile2 = FreeFile
    Open sFile2 For Input Access Read Lock Write As #lFile2
    
    Dim sText1 As String
    Dim sText2 As String
    Do While Not VBA.FileSystem.EOF(lFile1) And Not VBA.FileSystem.EOF(lFile2)
        Line Input #lFile1, sText1
        Line Input #lFile2, sText2
        If sText1 <> sText2 Then
            FilesEqual = False
            Exit Do
        End If
    Loop
    
    Close lFile2
    Close lFile1
End Function





Private Function SafeFileLen(ByVal sFile As String) As Long
    On Error Resume Next
    SafeFileLen = -1
    SafeFileLen = VBA.FileSystem.FileLen(sFile)
End Function






Private Function GetFileExtension(ByVal pComponent As VBComponent)
    Select Case pComponent.Type
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
            
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
            
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
            
        Case vbext_ct_Document
            ' Skip this type of module
            
        Case Else
            Debug.Assert False
    End Select
End Function





Private Sub SplitPath(ByVal sFullPath As String, _
    Optional ByRef sDir As String, _
    Optional ByRef sDriveOrShare As String, _
    Optional ByRef sPath As String, _
    Optional ByRef sFile As String, _
    Optional ByRef sFileTitle As String, _
    Optional ByRef sExtension As String)
    
    sDriveOrShare = ""
    sPath = ""
    sFile = ""
    sFileTitle = ""
    sExtension = ""
  
    Dim iPos As Long
    
    ' Determine drive or share:
    If sFullPath Like "?:*" Then ' sDriveOrShare:
        sDriveOrShare = Left$(sFullPath, 2)
        sFullPath = Mid$(sFullPath, 3)
    ElseIf sFullPath Like "\\*" Then ' \\Server
        iPos = InStr(3, sFullPath, "\")
        sDriveOrShare = Left$(sFullPath, iPos - 1)
        sFullPath = Mid$(sFullPath, iPos)
    End If
    
    ' Split path and file name:
    iPos = InStrRev(sFullPath, "\")
    sPath = Left$(sFullPath, iPos)
    sFile = Mid$(sFullPath, iPos + 1)
    
    sDir = sDriveOrShare & sPath
    
    ' Split file title and extension:
    iPos = InStrRev(sFile, ".")
    If iPos > 0 Then
        sFileTitle = Left$(sFile, iPos - 1)
        sExtension = Mid$(sFile, iPos)
    Else
        sFileTitle = sFile
    End If
End Sub














Private Sub Test_FilesEqual()
    Dim sDir As String
    sDir = BaseDir()
    
    ExportToCodeModules
    
    Debug.Assert FilesEqual(sDir & "Dev.bas", sDir & "Dev.bas")
    Debug.Assert Not FilesEqual(sDir & "Dev.bas", sDir & "Conversion.bas")
End Sub



Private Sub Test_SafeFileLen()
    Debug.Assert SafeFileLen("nul") = -1
    Debug.Assert SafeFileLen(Application.VBE.ActiveVBProject.FileName) > 0
End Sub



Private Sub Test_GetFileExtension()
    Debug.Assert GetFileExtension(Application.VBE.ActiveVBProject.VBComponents("Dev")) = ".bas"
End Sub



Private Sub Test_SplitPath()
    Dim sDir As String
    Dim sDriveOrShare As String
    Dim sPath As String
    Dim sFile As String
    Dim sFileTitle As String
    Dim sExtension As String
    
    Const PATH As String = "C:\This\Is\The\Path\To\My.file"
    
    SplitPath PATH, sDir, sDriveOrShare, sPath, sFile, sFileTitle, sExtension
    
    Debug.Assert sDriveOrShare = "C:"
    Debug.Assert sPath = "\This\Is\The\Path\To\"
    Debug.Assert sFile = "My.file"
    Debug.Assert sFileTitle = "My"
    Debug.Assert sExtension = ".file"
    
    SplitPath Application.VBE.ActiveVBProject.FileName, sDir, sDriveOrShare, sPath, sFile, sFileTitle, sExtension
    Debug.Assert sDriveOrShare & sPath & sFile = Application.VBE.ActiveVBProject.FileName
    Debug.Assert sDriveOrShare & sPath & sFileTitle & sExtension = Application.VBE.ActiveVBProject.FileName
    Debug.Assert sDir & sFile = Application.VBE.ActiveVBProject.FileName
    Debug.Assert sDir & sFileTitle & sExtension = Application.VBE.ActiveVBProject.FileName
End Sub
