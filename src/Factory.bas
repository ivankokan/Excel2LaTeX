Attribute VB_Name = "Factory"
Option Explicit

Public Function NewController() As IController
    Set NewController = New CController
End Function

Public Function NewModel() As IModel
    Set NewModel = New CModel
End Function

Public Function NewDefaultModel() As IModel
    Set NewDefaultModel = NewModel
    NewDefaultModel.InitDefault
End Function

Public Function NewView() As frmConvert
    Set NewView = New frmConvert
End Function

Function NewStorage() As IStorage
    Set NewStorage = New CSheetStorage
End Function


