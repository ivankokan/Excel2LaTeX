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

Public Sub IController_Latex(ByVal pView As frmConvert, ByVal pModel As IModel)
    If pModel.Range Is Nothing Then
        MsgBox "This macro converts the selected table to LaTeX. Please select a single rectangular range.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Load pView
    pView.Init Me, pModel
    pView.ConvertSelection
    pView.Show
    Exit Sub
End Sub

