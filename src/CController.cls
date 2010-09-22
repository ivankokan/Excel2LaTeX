VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IController

Public Event ModelChanged()

Public Sub Latex(ByVal pView As frmConvert, ByVal pModel As IModel)
    Dim selRange As Range
    If Selection Is Nothing Then GoTo ErrorMsg
    If TypeName(Selection) <> "Range" Then GoTo ErrorMsg
    If Selection.Areas.Count > 1 Then GoTo ErrorMsg
    
    Load pView
    Set selRange = Selection
    
    Dim sName As String
    On Error Resume Next
    sName = ActiveSheet.Name
    sName = selRange.Name.Name
    On Error GoTo 0
    
    pModel.FileName = CurDir + "\" + sName + ".tex"

    pView.Init Me, pModel
    pView.ConvertSelection
    pView.Show
    Exit Sub

ErrorMsg:
    MsgBox "This macro converts the selected table to LaTeX. Please select a single table", vbOKOnly + vbCritical
End Sub

