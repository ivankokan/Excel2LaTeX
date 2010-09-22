VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvert 
   Caption         =   "Exce2LaTeX"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   OleObjectBlob   =   "frmConvert.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mController As CController
Attribute mController.VB_VarHelpID = -1

Private mModel As IModel

Public Sub Init(ByVal pController As CController, ByVal pModel As IModel)
    Set mController = pController
    Set mModel = pModel
End Sub

Private Sub mController_ModelChanged()
'
End Sub

Public Sub ConvertSelection()
    InitModel mModel
    txtResult = mModel.GetConversionResult
    txtResult.SetFocus
End Sub

Public Sub InitModel(ByVal pModel As IModel)
    With pModel
        .CellWidth = Val(frmConvert.txtCellSize)
        .Options = Me.GetOptions()
        .Indent = Val(frmConvert.txtIndent)
    End With
End Sub

Public Sub InitFromModel(ByVal pModel As IModel)
    With pModel
        Me.txtCellSize = .CellWidth
        Me.SetOptions (.Options)
        Me.txtIndent = .Indent
    End With
End Sub

Function GetOptions() As x2lOptions
    If chkBooktabs.Value Then GetOptions = GetOptions Or x2lBooktabs
    If chkConvertDollar.Value Then GetOptions = GetOptions Or x2lConvertMathChars
    If chkTableFloat.Value Then GetOptions = GetOptions Or x2lCreateTableEnvironment
End Function
Sub SetOptions(ByVal Options As x2lOptions)
    chkBooktabs.Value = IIf(GetOptions And x2lBooktabs, stdole.Checked, stdole.Unchecked)
    chkConvertDollar.Value = IIf(GetOptions And x2lConvertMathChars, stdole.Checked, stdole.Unchecked)
    chkTableFloat.Value = IIf(GetOptions And x2lCreateTableEnvironment, stdole.Checked, stdole.Unchecked)
End Sub

Private Sub chkBooktabs_Click()
  ConvertSelection
End Sub

Private Sub chkConvertDollar_Click()
  ConvertSelection
End Sub

Private Sub chkTableFloat_Click()
  ConvertSelection
End Sub

Private Sub cmdBrowse_Click()
Dim FileName
  FileName = Application.GetSaveAsFilename(txtFilename, "TeX documents (*.tex), *.tex")
  If FileName <> False Then txtFilename = FileName
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub



Private Sub cmdCopy_Click()
  Dim dataObj As New DataObject
  dataObj.SetText txtResult
  dataObj.PutInClipboard
  Hide
End Sub

Private Sub cmdRefresh_Click()
  ConvertSelection
End Sub

Private Sub cmdSave_Click()
Dim FileName
  FileName = frmConvert.txtFilename
  If FileName = "" Then Exit Sub
  Open FileName For Output As 1
  Print #1, txtResult
  Close #1
  Hide
End Sub


Private Sub CommandButton2_Click()
  frmAbout.Show
End Sub

Private Sub spnCellWidth_Change()
  txtCellSize = spnCellWidth
End Sub

Private Sub spnIndent_Change()
  txtIndent = spnIndent
End Sub

Private Sub txtCellSize_Change()
  spnCellWidth = txtCellSize
  ConvertSelection
End Sub


Private Sub txtIndent_Change()
  spnIndent = txtIndent
  ConvertSelection
End Sub

Private Sub UserForm_Initialize()
  spnCellWidth = txtCellSize
End Sub

