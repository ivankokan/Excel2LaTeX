VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Model() As IModel
'
End Property
Public Property Set Model(ByVal pModel As IModel)
'
End Property

Public Property Get Controller() As IController
'
End Property
Public Property Set Controller(ByVal pController As IController)
'
End Property

Public Property Get Storage() As IStorage
'
End Property
Public Property Set Storage(ByVal pStorage As IStorage)
'
End Property

Public Sub Show(ByVal Modal As FormShowConstants)
'
End Sub
