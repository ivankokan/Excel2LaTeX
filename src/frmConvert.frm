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

