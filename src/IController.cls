VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Events() As IControllerEvents
'
End Property

Public Property Get Model() As IModel
'
End Property
Public Property Set Model(ByVal pModel As IModel)
'
End Property

Public Property Get View() As IView
'
End Property
Public Property Set View(ByVal pView As IView)
'
End Property

Public Property Get Storage() As IStorage
'
End Property
Public Property Set Storage(ByVal pStorage As IStorage)
'
End Property

Public Sub Run()
'
End Sub


