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

Public Sub IController_LaTeX(ByVal pView As IView, ByVal pModel As IModel, ByVal pStorage As IStorage)
    If pModel.RangeAddress = "" Then
        MsgBox "This macro converts the selected table to LaTeX. Please select a single rectangular range.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    With pView
        Set .Controller = Me
        Set .Model = pModel
        Set .Storage = pStorage
        .Show vbModeless
    End With
    
    Exit Sub
End Sub

