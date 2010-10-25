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

Implements IView

Private mController As IController
Attribute mController.VB_VarHelpID = -1

Private mModel As IModel
Private WithEvents mModelEvents As IModelEvents
Attribute mModelEvents.VB_VarHelpID = -1

Private mStorage As IStorage
Private WithEvents mStorageEvents As IStorageEvents
Attribute mStorageEvents.VB_VarHelpID = -1

Private WithEvents mActiveWkSheet As Worksheet
Attribute mActiveWkSheet.VB_VarHelpID = -1

Private mbIgnoreControlEvents As Boolean

'
' IView implementation
'
Private Property Get IView_Model() As IModel
    Set IView_Model = mModel
End Property
Private Property Set IView_Model(ByVal pModel As IModel)
    Set mModel = pModel
    Set mModelEvents = pModel.Events
    InitFromModel mModel
    
    Set mActiveWkSheet = Nothing
    If Not mModel.Range Is Nothing Then Set mActiveWkSheet = mModel.Range.Worksheet
End Property

Private Property Get IView_Controller() As IController
    Set IView_Controller = mController
End Property
Private Property Set IView_Controller(ByVal pController As IController)
    Set mController = pController
End Property

Private Property Get IView_Storage() As IStorage
    Set IView_Storage = mStorage
End Property
Private Property Set IView_Storage(ByVal pStorage As IStorage)
    Set mStorage = pStorage
    Set mStorageEvents = pStorage.Events
End Property

Private Sub IView_Show(ByVal Modal As FormShowConstants)
    Me.Show Modal
End Sub


'
' Form implementation
'
Private Function SafeRangePrecedents(ByVal pRange As Range) As Range
    On Error Resume Next
    Set SafeRangePrecedents = pRange.Precedents
End Function

Private Function UnionOfRangeAndItsPrecedents(ByVal pRange As Range) As Range
    Dim pPrecedents As Range
    Set pPrecedents = SafeRangePrecedents(pRange)
    
    If pPrecedents Is Nothing Then
        Set UnionOfRangeAndItsPrecedents = pRange
    Else
        Set UnionOfRangeAndItsPrecedents = Union(pRange, pPrecedents)
    End If
End Function

Private Sub mActiveWkSheet_Change(ByVal Target As Range)
    If Not Intersect(Target, UnionOfRangeAndItsPrecedents(mModel.Range)) Is Nothing Then
        ConvertSelection
    End If
End Sub

Private Sub mModelEvents_Changed()
    If mbIgnoreControlEvents Then Exit Sub
    txtResult = mModel.GetConversionResult
End Sub

Private Sub mStorageEvents_Changed()
'
End Sub

Private Sub ConvertSelection()
    If mbIgnoreControlEvents Then Exit Sub
    txtResult = mModel.GetConversionResult
    txtResult.SetFocus
End Sub

Public Sub InitModel(ByVal pModel As IModel)
    With pModel
        .CellWidth = Val(Me.txtCellSize)
        .Options = Me.GetOptions()
        .Indent = Val(Me.txtIndent)
        .FileName = Me.txtFilename
    End With
End Sub

Public Sub InitFromModel(ByVal pModel As IModel)
    mbIgnoreControlEvents = True
    With pModel
        Me.txtCellSize = .CellWidth
        Me.SetOptions (.Options)
        Me.txtIndent = .Indent
        Me.txtFilename = .FileName
    End With
    mbIgnoreControlEvents = False
    ConvertSelection
End Sub

Function GetOptions() As x2lOptions
    If chkBooktabs.Value Then GetOptions = GetOptions Or x2lBooktabs
    If chkConvertDollar.Value Then GetOptions = GetOptions Or x2lConvertMathChars
    If chkTableFloat.Value Then GetOptions = GetOptions Or x2lCreateTableEnvironment
End Function
Sub SetOptions(ByVal Options As x2lOptions)
    chkBooktabs.Value = (Options And x2lBooktabs) <> 0
    chkConvertDollar.Value = (Options And x2lConvertMathChars) <> 0
    chkTableFloat.Value = (Options And x2lCreateTableEnvironment) <> 0
End Sub

Private Sub UpdateOptions()
    mModel.Options = GetOptions()
End Sub

Private Sub chkBooktabs_Click()
    UpdateOptions
End Sub

Private Sub chkConvertDollar_Click()
    UpdateOptions
End Sub

Private Sub chkTableFloat_Click()
    UpdateOptions
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
    mModel.CellWidth = txtCellSize
End Sub

Private Sub txtIndent_Change()
    spnIndent = txtIndent
    mModel.Indent = txtIndent
End Sub


Private Sub UserForm_Click()
' This is regenerated every time the form is activated in the IDE. Just keep it here.
End Sub
