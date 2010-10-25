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

Private mView As IView
Private mModel As IModel
Private mStorage As IStorage

Private mpEvents As New IControllerEvents

Private Property Get IController_Events() As IControllerEvents
    Set IController_Events = mpEvents
End Property

Private Property Get IController_Model() As IModel
    Set IController_Model = mModel
End Property
Private Property Set IController_Model(ByVal pModel As IModel)
    Set mModel = pModel
    Set mView.Model = pModel
End Property

Private Property Get IController_View() As IView
    Set IController_View = mView
End Property
Private Property Set IController_View(ByVal pView As IView)
    Set mView = pView
    Set mView.Controller = Me
End Property

Private Property Get IController_Storage() As IStorage
    Set IController_Storage = mStorage
End Property
Private Property Set IController_Storage(ByVal pStorage As IStorage)
    Set mStorage = pStorage
    Set mView.Storage = pStorage
End Property

Private Sub IController_Run()
    mView.Show vbModeless
End Sub

