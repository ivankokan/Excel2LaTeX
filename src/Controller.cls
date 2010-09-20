VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Latex(ByVal pView As frmConvert)
    Dim sName As String
    Dim selRange As Range
    If Selection Is Nothing Then GoTo ErrorMsg
    If TypeName(Selection) <> "Range" Then GoTo ErrorMsg
    If Selection.Areas.Count > 1 Then GoTo ErrorMsg
    
    Load pView
    Set selRange = Selection
    On Error Resume Next
    sName = ActiveSheet.Name
    sName = selRange.Name.Name
    On Error GoTo 0
    
    pView.txtFilename = CurDir + "\" + sName + ".tex"
    pView.ConvertSelection
    pView.Show
    Exit Sub

ErrorMsg:
    MsgBox "This macro coverts the selected table to Latex. Pleas select a single table", vbOKOnly + vbCritical
End Sub

