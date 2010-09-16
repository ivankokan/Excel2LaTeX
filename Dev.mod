Attribute VB_Name = "Dev"
Option Explicit

Private Sub ExportToFiles()
    Dim sDir As String
    SplitPath Application.VBE.ActiveVBProject.FileName, sDir:=sDir
    
    Dim pVbComponent As VBComponent
    For Each pVbComponent In Application.VBE.ActiveVBProject.VBComponents
        pVbComponent.Export sDir & pVbComponent.name & GetFileExtension(pVbComponent)
    Next
End Sub







Public Function GetFileExtension(ByVal pComponent As VBComponent)
    Select Case pComponent.Type
        Case vbext_ct_StdModule
            GetFileExtension = ".mod"
            
        Case vbext_ct_Document, vbext_ct_ClassModule
            GetFileExtension = ".cls"
            
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
            
        Case Else
            Debug.Assert False
    End Select
End Function


Sub SplitPath(ByVal sFullPath As String, _
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
























Sub Test_SplitPath()
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
